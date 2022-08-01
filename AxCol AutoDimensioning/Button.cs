using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Media.Imaging;

namespace AxCol_AutoDimensioning
{
    internal class Button : IExternalApplication
    {
        public Result OnShutdown(UIControlledApplication application)
        {
            return Result.Succeeded;
        }

        public Result OnStartup(UIControlledApplication application)
        {
            // CREATING TAB
            application.CreateRibbonTab("ITI Plugins");

            // CREATING PANEL
            RibbonPanel ribPanel = application.CreateRibbonPanel("ITI Plugins", "Dimensioning");

            // CREATING BUTTON
            string assemblyName = Assembly.GetExecutingAssembly().Location;
            PushButtonData buttonData = new PushButtonData("AutoDim", "Auto\nDim", assemblyName, "AxCol_AutoDimensioning.AutoDimensioning")
            {
                ToolTip = "To auto dimension Axes and Columns all in the same moment.",
                LongDescription = "Get over-all dimensions between axes and columns for both near sides.",
            };

            // PUSHING THE BUTTON TO THE PANEL
            PushButton button = ribPanel.AddItem(buttonData) as PushButton;

            // SETTING AN ICON FOR THE BUTTON
            Uri uri = new Uri("pack://application:,,,/AxCol AutoDimensioning;component/Resources/C.png");
            button.LargeImage = new BitmapImage(uri);

            return Result.Succeeded;
        }
    }
}
