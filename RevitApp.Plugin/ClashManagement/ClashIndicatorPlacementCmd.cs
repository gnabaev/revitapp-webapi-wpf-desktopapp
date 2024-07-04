using AngleSharp.Html.Dom;
using AngleSharp.Html.Parser;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Microsoft.Win32;
using System.IO;
using System.Linq;
using System.Text;

namespace RevitApp.Plugin.ClashManagement
{
    [Transaction(TransactionMode.Manual)]
    public class ClashIndicatorPlacementCmd : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            var fileNames = GetFileNamesFromDialog("Открыть HTML файл", "HTML Files (*.html)|*.html", false);

            if (fileNames.Length > 0)
            {
                Document doc = commandData.Application.ActiveUIDocument.Document;

                // Get coordinates of the document base point
                var docBasePoint = BasePoint.GetProjectBasePoint(doc).SharedPosition;
                var docBasePointX = docBasePoint.X;
                var docBasePointY = docBasePoint.Y;
                var docBasePointZ = docBasePoint.Z;

                // Get the correct document title if the user is using a local copy of the central model
                var docTitle = GetDocumentTitle(doc);

                // Get a clash indicator family symbol
                var indicatorSymbol = new FilteredElementCollector(doc)
                    .OfClass(typeof(FamilySymbol))
                    .Cast<FamilySymbol>()
                    .FirstOrDefault(f => f.FamilyName == "Индикатор коллизии");

                if (indicatorSymbol == null)
                {
                    TaskDialog.Show("Ошибка", $"В документе \"{docTitle}\" отсутствует семейство \"Индикатор коллизии\". Загрузите семейство и повторите попытку.");
                    return Result.Cancelled;
                }

                foreach (var fileName in fileNames)
                {
                    //Get HTML document tree
                    var htmlDoc = GetHtmlDocument(fileName);

                    var reportName = htmlDoc.QuerySelector(".testName").InnerHtml;

                    // Get the main table with clashes data
                    var mainTable = htmlDoc.QuerySelector(".mainTable");

                    var mainTableSection = mainTable.Children.FirstOrDefault();

                    // Get all rows of the main table
                    var mainTableRows = mainTableSection.Children;

                    var headerRows = mainTableRows.Where(i => i.ClassName == "headerRow").ToList();

                    var headerRow = headerRows[1];

                    var headerColumns = headerRow.Children;

                    var generalHeaderColumns = headerColumns.Where(c => c.ClassName == "generalHeader").ToList();

                    int clashIndex = -1;

                    int pointIndex = -1;

                    for (int i = 0; i < generalHeaderColumns.Count; i++)
                    {
                        if (generalHeaderColumns[i].InnerHtml == "Наименование конфликта")
                        {
                            clashIndex = i;
                        }

                        if (generalHeaderColumns[i].InnerHtml == "Точка конфликта")
                        {
                            pointIndex = i;
                        }
                    }

                    var itemHeaderColumns = headerColumns.Where(c => c.ClassName == "item1Header").ToList();

                    int itemIdIndex = -1;

                    int itemModelIndex = -1;

                    for (int i = 0; i < itemHeaderColumns.Count; i++)
                    {
                        if (itemHeaderColumns[i].InnerHtml == "Id")
                        {
                            itemIdIndex = i;
                        }

                        if (itemHeaderColumns[i].InnerHtml == "Файл источника")
                        {
                            itemModelIndex = i;
                        }
                    }

                    var contentRows = mainTableRows.Where(i => i.ClassName == "contentRow").ToList();

                    if (contentRows.Count == 0)
                    {
                        TaskDialog.Show("Предупреждение", $"В отчете \"{reportName}\" коллизий не обнаружено. Данный отчет будет пропущен.");
                        continue;
                    }

                    foreach (var contentRow in contentRows)
                    {
                        var contentColumns = contentRow.Children;

                        var clashName = contentColumns[clashIndex].InnerHtml;

                        var clashPointCoordinates = contentColumns[pointIndex].InnerHtml;

                        var clashPoint = GetClashPoint(clashPointCoordinates, docBasePointX, docBasePointY, docBasePointZ);

                        var element1ContentColumns = contentColumns.Where(c => c.ClassName == "элемент1Содержимое").ToList();

                        var clashElementId1 = new ElementId(int.Parse(element1ContentColumns[itemIdIndex].InnerHtml));

                        var modelName1 = element1ContentColumns[itemModelIndex].InnerHtml;

                        var element2ContentColumns = contentColumns.Where(c => c.ClassName == "элемент2Содержимое").ToList();

                        var clashElementId2 = new ElementId(int.Parse(element2ContentColumns[itemIdIndex].InnerHtml));

                        var modelName2 = element2ContentColumns[itemModelIndex].InnerHtml;

                        if (modelName1 == docTitle && modelName2 == docTitle)
                        {
                            var element1 = doc.GetElement(clashElementId1);
                            var element2 = doc.GetElement(clashElementId2);

                            if (element1 != null && element2 != null)
                            {
                                using (Transaction transaction = new Transaction(doc))
                                {
                                    transaction.Start("Размещение индикатора коллизии");

                                    var indicatorInstance = PlaceClashIndicator(doc, clashPoint, indicatorSymbol);

                                    FillClashIndicatorInfo(indicatorInstance, reportName, clashName, clashElementId1, modelName1, clashElementId2, modelName2);

                                    transaction.Commit();
                                }
                            }
                        }

                        else if (modelName1 != modelName2 && modelName1 == docTitle)
                        {
                            var rvtLinks = new FilteredElementCollector(doc).OfClass(typeof(RevitLinkInstance)).Cast<RevitLinkInstance>().ToList();

                            if (rvtLinks.Count == 0)
                            {
                                TaskDialog.Show("Ошибка", $"В документе \"{docTitle}\" отсутствуют RVT-связи. Загрузите минимум одну RVT-связь для размещения индикаторов коллизий.");
                                return Result.Cancelled;
                            }

                            Document linkDoc;

                            try
                            {
                                linkDoc = rvtLinks.Select(l => l.GetLinkDocument()).FirstOrDefault(ld => modelName2 == ld.Title + ".rvt");
                            }
                            catch (System.NullReferenceException)
                            {
                                continue;
                            }

                            var element1 = doc.GetElement(clashElementId1);
                            var element2 = linkDoc.GetElement(clashElementId2);

                            if (element1 != null && element2 != null)
                            {
                                using (Transaction transaction = new Transaction(doc))
                                {
                                    transaction.Start("Размещение индикатора коллизии");

                                    var indicatorInstance = PlaceClashIndicator(doc, clashPoint, indicatorSymbol);

                                    FillClashIndicatorInfo(indicatorInstance, reportName, clashName, clashElementId1, modelName1, clashElementId2, modelName2);

                                    transaction.Commit();
                                }
                            }
                        }

                        else if (modelName1 != modelName2 && modelName2 == docTitle)
                        {
                            var rvtLinks = new FilteredElementCollector(doc).OfClass(typeof(RevitLinkInstance)).Cast<RevitLinkInstance>().ToList();

                            if (rvtLinks.Count == 0)
                            {
                                TaskDialog.Show("Ошибка", $"В документе \"{docTitle}\" отсутствуют RVT-связи. Загрузите минимум одну RVT-связь для размещения индикаторов коллизий.");
                                return Result.Cancelled;
                            }

                            Document linkDoc;

                            try
                            {
                                linkDoc = rvtLinks.Select(l => l.GetLinkDocument()).FirstOrDefault(ld => modelName1 == ld.Title + ".rvt");
                            }
                            catch (System.NullReferenceException)
                            {
                                continue;
                            }

                            var element1 = linkDoc.GetElement(clashElementId1);
                            var element2 = doc.GetElement(clashElementId2);

                            if (element1 != null && element2 != null)
                            {
                                using (Transaction transaction = new Transaction(doc))
                                {
                                    transaction.Start("Размещение индикатора коллизии");

                                    var indicatorInstance = PlaceClashIndicator(doc, clashPoint, indicatorSymbol);

                                    FillClashIndicatorInfo(indicatorInstance, reportName, clashName, clashElementId1, modelName1, clashElementId2, modelName2);

                                    transaction.Commit();
                                }
                            }
                        }

                        else
                        {
                            continue;
                        }
                    }
                }
            }

            else
            {
                return Result.Cancelled;
            }

            return Result.Succeeded;
        }

        private string[] GetFileNamesFromDialog(string title, string filter, bool multiselect)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();

            openFileDialog.Title = title;
            openFileDialog.Filter = filter;
            openFileDialog.Multiselect = multiselect;

            bool? result = openFileDialog.ShowDialog();

            if (result == true)
            {
                return openFileDialog.FileNames;
            }
            else
            {
                return new string[0];
            }
        }

        private string GetDocumentTitle(Document doc)
        {
            string docTitle = doc.Title;

            if (!doc.IsWorkshared)
            {
                docTitle += ".rvt";
            }
            else
            {
                var separatorIndex = docTitle.LastIndexOf('_');

                if (separatorIndex != -1)
                {
                    docTitle = docTitle.Substring(0, separatorIndex) + ".rvt";
                }
            }

            return docTitle;
        }

        private IHtmlDocument GetHtmlDocument(string fileName)
        {
            string htmlFile = File.ReadAllText(fileName, Encoding.UTF8);

            IHtmlDocument htmlDoc = new HtmlParser().ParseDocument(htmlFile);

            return htmlDoc;
        }

        private XYZ GetClashPoint(string clashPointCoordinates, double basePointX, double basePointY, double basePointZ)
        {
            var pointCoordinates = clashPointCoordinates.Split(',').SelectMany(x => x.Trim().Split(':')).Where((x, i) => i % 2 != 0).Select(x => UnitUtils.Convert(double.Parse(x.Replace('.', ',')), UnitTypeId.Meters, UnitTypeId.Feet)).ToArray();

            var clashPoint = new XYZ(pointCoordinates[0] - basePointX, pointCoordinates[1] - basePointY, pointCoordinates[2] - basePointZ);

            return clashPoint;
        }

        private FamilyInstance PlaceClashIndicator(Document doc, XYZ point, FamilySymbol familySymbol)
        {
            if (!familySymbol.IsActive)
            {
                familySymbol.Activate();
            }

            var familyInstance = doc.Create.NewFamilyInstance(point, familySymbol, Autodesk.Revit.DB.Structure.StructuralType.NonStructural);
            familyInstance.Pinned = true;

            return familyInstance;
        }

        private void FillClashIndicatorInfo(FamilyInstance clashIndicator, string reportName, string clashName, ElementId clashElementId1, string modelName1, ElementId clashElementId2, string modelName2)
        {
            if (clashIndicator != null)
            {
                clashIndicator.LookupParameter("V Наименование отчета").Set(reportName);
                clashIndicator.LookupParameter("V Наименование конфликта").Set(clashName);

                clashIndicator.LookupParameter("V Идентификатор 1").Set(clashElementId1.ToString());
                clashIndicator.LookupParameter("V Модель 1").Set(modelName1);

                clashIndicator.LookupParameter("V Идентификатор 2").Set(clashElementId2.ToString());
                clashIndicator.LookupParameter("V Модель 2").Set(modelName2);
            }
        }
    }
}
