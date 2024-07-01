using AngleSharp.Html.Dom;
using AngleSharp.Html.Parser;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.Exceptions;
using Autodesk.Revit.UI;
using Microsoft.Win32;
using System.Collections.Generic;
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
            Document doc = commandData.Application.ActiveUIDocument.Document;

            string docTitle;

            if (!doc.IsWorkshared)
            {
                docTitle = doc.Title;
            }
            else
            {
                int separatorIndex = doc.Title.LastIndexOf('_');

                if (separatorIndex != -1)
                {
                    docTitle = doc.Title.Substring(0, separatorIndex);
                }
                else
                {
                    return Result.Cancelled;
                }
            }

            FamilySymbol indicatorSymbol = new FilteredElementCollector(doc)
                .OfClass(typeof(FamilySymbol))
                .Cast<FamilySymbol>()
                .FirstOrDefault(f => f.FamilyName == "Индикатор коллизии");

            if (indicatorSymbol == null)
            {
                TaskDialog.Show("Ошибка", $"В документе \"{docTitle}\" отсутствует семейство \"Индикатор коллизии\". Загрузите семейство и повторите попытку.");
                return Result.Cancelled;
            }

            var rvtLinks = new FilteredElementCollector(doc).OfClass(typeof(RevitLinkInstance)).Cast<RevitLinkInstance>().ToList();

            if (rvtLinks.Count == 0)
            {
                TaskDialog.Show("Ошибка", $"В документе \"{docTitle}\" отсутствуют RVT-связи. Загрузите минимум одну RVT-связь для размещения индикаторов коллизий.");
                return Result.Cancelled;
            }

            OpenFileDialog openFileDialog = new OpenFileDialog();

            openFileDialog.Filter = "HTML Files (*.html)|*.html";
            openFileDialog.Multiselect = false;
            openFileDialog.Title = "Открыть HTML файл";

            bool? result = openFileDialog.ShowDialog();

            if (result == true)
            {
                foreach (var fileName in openFileDialog.FileNames)
                {
                    var htmlFile = File.ReadAllText(fileName, Encoding.UTF8);

                    //Get HTML document tree
                    IHtmlDocument htmlDoc = new HtmlParser().ParseDocument(htmlFile);

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
                        TaskDialog.Show("Ошибка", $"В отчете \"{reportName}\" коллизий не обнаружено.");
                        continue;
                    }

                    foreach (var contentRow in contentRows)
                    {
                        var contentColumns = contentRow.Children;

                        var clashName = contentColumns[clashIndex].InnerHtml;

                        var clashPoint = contentColumns[pointIndex].InnerHtml;

                        var pointCoordinates = clashPoint.Split(',').SelectMany(x => x.Trim().Split(':')).Where((x, i) => i % 2 != 0).Select(x => UnitUtils.Convert(double.Parse(x.Replace('.', ',')), UnitTypeId.Meters, UnitTypeId.Feet)).ToArray();

                        var point = new XYZ(pointCoordinates[0], pointCoordinates[1], pointCoordinates[2]);

                        var element1ContentColumns = contentColumns.Where(c => c.ClassName == "элемент1Содержимое").ToList();

                        int clashElementStringId1 = int.Parse(element1ContentColumns[itemIdIndex].InnerHtml);

                        string modelName1 = element1ContentColumns[itemModelIndex].InnerHtml;

                        var clashElement1 = new ClashElement(clashName, clashElementStringId1, modelName1);

                        var element2ContentColumns = contentColumns.Where(c => c.ClassName == "элемент2Содержимое").ToList();

                        int clashElementStringId2 = int.Parse(element2ContentColumns[itemIdIndex].InnerHtml);

                        string modelName2 = element2ContentColumns[itemModelIndex].InnerHtml;

                        var clashElement2 = new ClashElement(clashName, clashElementStringId2, modelName2);

                        var clashElementId1 = new ElementId(clashElement1.Id);
                        var clashElementId2 = new ElementId(clashElement2.Id);

                        if (modelName1.Contains(docTitle) && modelName2.Contains(docTitle))
                        {
                            var element1 = doc.GetElement(clashElementId1);
                            var element2 = doc.GetElement(clashElementId2);

                            if (element1 != null && element2 != null)
                            {
                                using (Transaction transaction = new Transaction(doc))
                                {
                                    transaction.Start("Размещение индикатора коллизии");

                                    if (!indicatorSymbol.IsActive)
                                    {
                                        indicatorSymbol.Activate();
                                    }

                                    var indicatorInstance = doc.Create.NewFamilyInstance(point, indicatorSymbol, Autodesk.Revit.DB.Structure.StructuralType.NonStructural);
                                    indicatorInstance.Pinned = true;

                                    FillClashIndicatorInfo(indicatorInstance, reportName, clashName, clashElementId1, modelName1, clashElementId2, modelName2);

                                    transaction.Commit();
                                }
                            }
                        }

                        else if (modelName1 != modelName2 && modelName1.Contains(docTitle))
                        {
                            Document linkDoc;

                            try
                            {
                                linkDoc = rvtLinks.Select(l => l.GetLinkDocument()).FirstOrDefault(d => modelName2.Contains(d.Title));
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

                                    if (!indicatorSymbol.IsActive)
                                    {
                                        indicatorSymbol.Activate();
                                    }

                                    var indicatorInstance = doc.Create.NewFamilyInstance(point, indicatorSymbol, Autodesk.Revit.DB.Structure.StructuralType.NonStructural);
                                    indicatorInstance.Pinned = true;

                                    FillClashIndicatorInfo(indicatorInstance, reportName, clashName, clashElementId1, modelName1, clashElementId2, modelName2);

                                    transaction.Commit();
                                }
                            }
                        }

                        else if (modelName1 != modelName2 && modelName2.Contains(docTitle))
                        {
                            Document linkDoc;

                            try
                            {
                                linkDoc = rvtLinks.Select(l => l.GetLinkDocument()).FirstOrDefault(d => modelName1.Contains(d.Title));
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

                                    if (!indicatorSymbol.IsActive)
                                    {
                                        indicatorSymbol.Activate();
                                    }

                                    var indicatorInstance = doc.Create.NewFamilyInstance(point, indicatorSymbol, Autodesk.Revit.DB.Structure.StructuralType.NonStructural);
                                    indicatorInstance.Pinned = true;

                                    FillClashIndicatorInfo(indicatorInstance, reportName, clashName, clashElementId1, modelName1, clashElementId2, modelName2);

                                    transaction.Commit();
                                }
                            }
                        }

                        else
                        {
                            TaskDialog.Show("Предупреждение", $"Выбран неверный отчет \"{reportName}\". По всем коллизиям обнаружено несоответствие наименований моделей наименованию текущего документа \"{docTitle}\" и связанных моделей. Данный отчет будет пропущен.");
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

        private Solid GetSolid(Element element)
        {
            List<Solid> solids = new List<Solid>();

            Options options = new Options();

            options.DetailLevel = ViewDetailLevel.Fine;
            options.ComputeReferences = false;
            options.IncludeNonVisibleObjects = false;

            GeometryElement geomElement = element.get_Geometry(options);

            foreach (GeometryObject geomObject in geomElement)
            {
                if (geomObject is Solid)
                {
                    var solid = geomObject as Solid;

                    if (solid.Volume > 0)
                    {
                        solids.Add(solid);
                    }
                }

                else if (geomObject is GeometryInstance)
                {
                    GeometryInstance geomInstance = geomObject as GeometryInstance;

                    GeometryElement instanceGeomElement = geomInstance.GetInstanceGeometry();

                    foreach (GeometryObject instanceGeomObject in instanceGeomElement)
                    {
                        if (instanceGeomObject is Solid)
                        {
                            var solid = instanceGeomObject as Solid;

                            if (solid.Volume > 0)
                            {
                                solids.Add(solid);
                            }
                        }
                    }
                }
            }

            Solid result = GetCombinedSolid(solids);

            return result;
        }

        private Solid GetCombinedSolid(IEnumerable<Solid> solids)
        {
            Solid combinedSolid;

            if (solids.Count() == 1)
            {
                combinedSolid = solids.FirstOrDefault();
            }
            else
            {
                combinedSolid = solids.FirstOrDefault();

                for (int i = 1; i < solids.Count(); i++)
                {
                    try
                    {
                        // Try to combine solids to one solid. If the current solid can't be combined with the previous solid, then it is skipped.
                        combinedSolid = BooleanOperationsUtils.ExecuteBooleanOperation(combinedSolid, solids.ElementAt(i), BooleanOperationsType.Union);
                    }
                    catch (InvalidOperationException)
                    {
                        continue;
                    }
                }
            }

            return combinedSolid;
        }
    }
}
