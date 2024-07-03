using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;

namespace RevitApp.Plugin.ElementIdUpdate
{
    [Transaction(TransactionMode.Manual)]
    public class ElementIdUpdateCmd : IExternalCommand
    {
        private ElementIdUpdater updater = null;

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            Application app = uiapp.Application;

            if (updater == null)
            {
                // Register updater to react to element creation
                updater = new ElementIdUpdater(app.ActiveAddInId);

                UpdaterRegistry.RegisterUpdater(updater);

                List<Type> classes = new List<Type>()
                {
                    typeof(Pipe),
                    typeof(PipeInsulation),
                    typeof(FlexPipe),
                    typeof(Duct),
                    typeof(DuctInsulation),
                    typeof(DuctLining),
                    typeof(FlexDuct),
                    typeof(FamilyInstance)
                };

                ElementMulticlassFilter multiclassFilter = new ElementMulticlassFilter(classes);

                UpdaterRegistry.AddTrigger(updater.GetUpdaterId(), multiclassFilter, Element.GetChangeTypeElementAddition());
            }
            else
            {
                UpdaterRegistry.UnregisterUpdater(updater.GetUpdaterId());

                updater = null;
            }

            return Result.Succeeded;
        }
    }
}
