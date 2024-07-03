using Autodesk.Revit.UI;
using System.Reflection;
using System;
using System.IO;
using RevitApp.Plugin.ClashManagement;
using System.Windows.Media.Imaging;

namespace RevitApp.Plugin
{
    public class ApplicationUI : IExternalApplication
    {
        public Result OnShutdown(UIControlledApplication application)
        {
            return Result.Succeeded;
        }

        public Result OnStartup(UIControlledApplication application)
        {
            string assemblyLocation = Assembly.GetExecutingAssembly().Location;
            string iconsDirectoryPath = Path.GetDirectoryName(assemblyLocation) + @"\icons\";

            string tabName = "WildTeam";
            application.CreateRibbonTab(tabName);

            RibbonPanel clashManagementPanel = application.CreateRibbonPanel(tabName, "Управление коллизиями");

            PushButtonData clashIndicatorPlacementButton = new PushButtonData(nameof(ClashIndicatorPlacementCmd), "Размещение индикатора", assemblyLocation, typeof(ClashIndicatorPlacementCmd).FullName)
            {
                LargeImage = new BitmapImage(new Uri(iconsDirectoryPath + "ClashIndicator.png"))
            };

            clashManagementPanel.AddItem(clashIndicatorPlacementButton);

            return Result.Succeeded;
        }
    }
}
