using AngleSharp.Html.Dom;
using AngleSharp.Html.Parser;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Microsoft.Win32;
using System;
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
            openFileDialog.Multiselect = true;
            openFileDialog.Title = "Открыть HTML файлы";

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

                    int clashIndex = 0;

                    for (int i = 0; i < generalHeaderColumns.Count; i++)
                    {
                        if (generalHeaderColumns[i].InnerHtml == "Наименование конфликта")
                        {
                            clashIndex = i;
                            break;
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

                    Dictionary<string, List<ClashElement>> clashes = new Dictionary<string, List<ClashElement>>();

                    var contentRows = mainTableRows.Where(i => i.ClassName == "contentRow").ToList();

                    if (contentRows.Count == 0)
                    {
                        TaskDialog.Show("Ошибка", $"В отчете \"{reportName}\" коллизий не обнаружено.");
                        continue;
                    }

                    foreach (var contentRow in contentRows)
                    {
                        var clashElements = new List<ClashElement>();

                        var contentColumns = contentRow.Children;

                        var clashName = contentColumns[clashIndex].InnerHtml;

                        var element1ContentColumns = contentColumns.Where(c => c.ClassName == "элемент1Содержимое").ToList();

                        int clashElementId1 = int.Parse(element1ContentColumns[itemIdIndex].InnerHtml);

                        string modelName1 = element1ContentColumns[itemModelIndex].InnerHtml;

                        var clashElement1 = new ClashElement(clashName, clashElementId1, modelName1);

                        clashElements.Add(clashElement1);

                        var element2ContentColumns = contentColumns.Where(c => c.ClassName == "элемент2Содержимое").ToList();

                        int clashElemenId2 = int.Parse(element2ContentColumns[itemIdIndex].InnerHtml);

                        string modelName2 = element2ContentColumns[itemModelIndex].InnerHtml;

                        var clashElement2 = new ClashElement(clashName, clashElemenId2, modelName2);

                        clashElements.Add(clashElement2);

                        clashes.Add(clashName, clashElements);
                    }

                    foreach (var conflict in clashes)
                    {
                        var clashName = conflict.Key;

                        var clashElement1 = conflict.Value[0];
                        var clashElement2 = conflict.Value[1];

                        var modelName1 = clashElement1.Model;
                        var modelName2 = clashElement2.Model;

                        var clashElementId1 = new ElementId(clashElement1.Id);
                        var clashElementId2 = new ElementId(clashElement2.Id);

                        if (modelName1.Contains(docTitle) && modelName2.Contains(docTitle))
                        {
                            var element1 = doc.GetElement(clashElementId1);
                            var element2 = doc.GetElement(clashElementId2);

                            if (element1 != null && element2 != null) 
                            {
                                Options options = new Options();
                                options.ComputeReferences = false;
                                options.IncludeNonVisibleObjects = false;

                                GeometryElement geometryElement1 = element1.get_Geometry(options);
                                GeometryElement geometryElement2 = element2.get_Geometry(options);

                                foreach (GeometryObject geometryObject1 in geometryElement1)
                                {
                                    foreach (GeometryObject geometryObject2 in geometryElement2)
                                    {
                                        // Transform both geometry objects to solid
                                        Solid solid1 = geometryObject1 as Solid;
                                        Solid solid2 = geometryObject2 as Solid;

                                        if (solid1 != null && solid2 != null)
                                        {
                                            // Perform a boolean operation to find the intersection
                                            Solid intersectionSolid = BooleanOperationsUtils.ExecuteBooleanOperation(solid1, solid2, BooleanOperationsType.Intersect);

                                            if (intersectionSolid != null && intersectionSolid.Volume > 0)
                                            {
                                                XYZ centroid = intersectionSolid.ComputeCentroid();

                                                using (Transaction transaction = new Transaction(doc))
                                                {
                                                    transaction.Start("Indicator placement");

                                                    if (!indicatorSymbol.IsActive)
                                                    {
                                                        indicatorSymbol.Activate();
                                                    }

                                                    var indicatorInstance = doc.Create.NewFamilyInstance(centroid, indicatorSymbol, Autodesk.Revit.DB.Structure.StructuralType.NonStructural);
                                                    indicatorInstance.Pinned = true;

                                                    indicatorInstance.LookupParameter("V Наименование отчета").Set(reportName);
                                                    indicatorInstance.LookupParameter("V Наименование конфликта").Set(clashName);

                                                    indicatorInstance.LookupParameter("V Идентификатор 1").Set(clashElementId1.ToString());
                                                    indicatorInstance.LookupParameter("V Модель 1").Set(modelName1);

                                                    indicatorInstance.LookupParameter("V Идентификатор 2").Set(clashElementId2.ToString());
                                                    indicatorInstance.LookupParameter("V Модель 2").Set(modelName2);

                                                    transaction.Commit();
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }

                        else if (modelName1 != modelName2 && modelName1.Contains(docTitle))
                        {
                            var linkDocs = rvtLinks.Select(l => l.GetLinkDocument()).Where(d => d.Title.Contains(modelName2)).ToList();

                            if (linkDocs.Count > 0)
                            {
                                foreach (var linkDoc in linkDocs)
                                {
                                    var element1 = doc.GetElement(clashElementId1);
                                    var element2 = linkDoc.GetElement(clashElementId2);

                                    if (element1 != null && element2 != null)
                                    {
                                        Options options = new Options();
                                        options.ComputeReferences = false;
                                        options.IncludeNonVisibleObjects = false;

                                        GeometryElement geometryElement1 = element1.get_Geometry(options);
                                        GeometryElement geometryElement2 = element2.get_Geometry(options);

                                        foreach (GeometryObject geometryObject1 in geometryElement1)
                                        {
                                            foreach (GeometryObject geometryObject2 in geometryElement2)
                                            {
                                                // Transform both geometry objects to solid
                                                Solid solid1 = geometryObject1 as Solid;
                                                Solid solid2 = geometryObject2 as Solid;

                                                if (solid1 != null && solid2 != null)
                                                {
                                                    // Perform a boolean operation to find the intersection
                                                    Solid intersectionSolid = BooleanOperationsUtils.ExecuteBooleanOperation(solid1, solid2, BooleanOperationsType.Intersect);

                                                    if (intersectionSolid != null && intersectionSolid.Volume > 0)
                                                    {
                                                        XYZ centroid = intersectionSolid.ComputeCentroid();

                                                        using (Transaction transaction = new Transaction(doc))
                                                        {
                                                            transaction.Start("Indicator placement");

                                                            if (!indicatorSymbol.IsActive)
                                                            {
                                                                indicatorSymbol.Activate();
                                                            }

                                                            var indicatorInstance = doc.Create.NewFamilyInstance(centroid, indicatorSymbol, Autodesk.Revit.DB.Structure.StructuralType.NonStructural);
                                                            indicatorInstance.Pinned = true;

                                                            indicatorInstance.LookupParameter("V Наименование отчета").Set(reportName);
                                                            indicatorInstance.LookupParameter("V Наименование конфликта").Set(clashName);

                                                            indicatorInstance.LookupParameter("V Идентификатор 1").Set(clashElementId1.ToString());
                                                            indicatorInstance.LookupParameter("V Модель 1").Set(modelName1);

                                                            indicatorInstance.LookupParameter("V Идентификатор 2").Set(clashElementId2.ToString());
                                                            indicatorInstance.LookupParameter("V Модель 2").Set(modelName2);

                                                            transaction.Commit();
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }

                        else if (modelName1 != modelName2 && modelName2.Contains(docTitle))
                        {
                            var linkDocs = rvtLinks.Select(l => l.GetLinkDocument()).Where(d => d.Title.Contains(modelName1)).ToList();

                            if (linkDocs.Count > 0)
                            {
                                foreach (var linkDoc in linkDocs)
                                {
                                    var element1 = linkDoc.GetElement(clashElementId1);
                                    var element2 = doc.GetElement(clashElementId2);

                                    if (element1 != null && element2 != null)
                                    {
                                        Options options = new Options();
                                        options.ComputeReferences = false;
                                        options.IncludeNonVisibleObjects = false;

                                        GeometryElement geometryElement1 = element1.get_Geometry(options);
                                        GeometryElement geometryElement2 = element2.get_Geometry(options);

                                        foreach (GeometryObject geometryObject1 in geometryElement1)
                                        {
                                            foreach (GeometryObject geometryObject2 in geometryElement2)
                                            {
                                                // Transform both geometry objects to solid
                                                Solid solid1 = geometryObject1 as Solid;
                                                Solid solid2 = geometryObject2 as Solid;

                                                if (solid1 != null && solid2 != null)
                                                {
                                                    // Perform a boolean operation to find the intersection
                                                    Solid intersectionSolid = BooleanOperationsUtils.ExecuteBooleanOperation(solid1, solid2, BooleanOperationsType.Intersect);

                                                    if (intersectionSolid != null && intersectionSolid.Volume > 0)
                                                    {
                                                        XYZ centroid = intersectionSolid.ComputeCentroid();

                                                        using (Transaction transaction = new Transaction(doc))
                                                        {
                                                            transaction.Start("Indicator placement");

                                                            if (!indicatorSymbol.IsActive)
                                                            {
                                                                indicatorSymbol.Activate();
                                                            }

                                                            var indicatorInstance = doc.Create.NewFamilyInstance(centroid, indicatorSymbol, Autodesk.Revit.DB.Structure.StructuralType.NonStructural);
                                                            indicatorInstance.Pinned = true;

                                                            indicatorInstance.LookupParameter("V Наименование отчета").Set(reportName);
                                                            indicatorInstance.LookupParameter("V Наименование конфликта").Set(clashName);

                                                            indicatorInstance.LookupParameter("V Идентификатор 1").Set(clashElementId1.ToString());
                                                            indicatorInstance.LookupParameter("V Модель 1").Set(modelName1);

                                                            indicatorInstance.LookupParameter("V Идентификатор 2").Set(clashElementId2.ToString());
                                                            indicatorInstance.LookupParameter("V Модель 2").Set(modelName2);

                                                            transaction.Commit();
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }

                        else
                        {
                            TaskDialog.Show("Предупреждение", $"Выбран неверный отчет \"{reportName}\". В нем обнаружено несоответствие наименований моделей наименованию текущего документа \"{docTitle}\" и связанных моделей. Данный отчет будет пропущен.");
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
    }
}
