using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;

namespace RevitApp.Plugin.ElementIdUpdate
{
    public class ElementIdUpdater : IUpdater
    {
        private AddInId _addInId = null; 
        private UpdaterId _updaterId = null;

        public ElementIdUpdater(AddInId addInId)
        {
            _addInId = addInId;
            _updaterId = new UpdaterId(_addInId, new Guid("56c0b7e1-3ae5-4cf7-99fa-ac3c730f8e05"));
        }

        public void Execute(UpdaterData data)
        {
            try
            {
                Document doc = data.GetDocument();

                foreach (ElementId id in data.GetAddedElementIds())
                {
                    Element element = doc.GetElement(id);

                    if (element != null)
                    {
                        element.LookupParameter("II Идентификатор").Set(id.IntegerValue);
                    }

                    doc.Regenerate();
                }
            }
            catch (Exception ex) 
            {
                TaskDialog.Show("Ошибка", ex.ToString());
            }

            return;
        }

        public string GetAdditionalInformation()
        {
            return "Automatic update element Id";
        }

        public ChangePriority GetChangePriority()
        {
            return ChangePriority.FreeStandingComponents;
        }

        public UpdaterId GetUpdaterId()
        {
            return _updaterId;
        }

        public string GetUpdaterName()
        {
            return "Element Id Updater";
        }
    }
}
