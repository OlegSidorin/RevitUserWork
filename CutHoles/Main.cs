using Autodesk.Revit.UI;
using Autodesk.Revit.DB;
using System.IO;
using System.Reflection;
using System.Windows.Media.Imaging;
using Autodesk.Revit.Attributes;
using System;
using Autodesk.Revit.DB.Events;
using System.Collections.ObjectModel;
using System.Collections.Generic;
using Autodesk.Revit.ApplicationServices;

namespace CutHoles
{
    [Transaction(TransactionMode.Manual), Regeneration(RegenerationOption.Manual)]
    class Main : IExternalApplication
    {
        public static string DllLocation { get; set; }
        public static string DllFolderLocation { get; set; }
        public static string UserFolder { get; set; }
        public static string TabName { get; set; } = "Надстройки";
        public static string PanelName { get; set; } = "M1";
        private static RibbonPanel Panel { get; set; }
        public Result OnStartup(UIControlledApplication application)
        {
            bool createPanelM1 = true;
            var ribbonPanels = application.GetRibbonPanels();
            foreach (var rp in ribbonPanels)
            {
                if (rp.Name == "M1") createPanelM1 = false;
                Panel = rp;
            }
            if (createPanelM1)
            {
                Panel = application.CreateRibbonPanel(PanelName);
            }
             
            DllLocation = Assembly.GetExecutingAssembly().Location;
            DllFolderLocation = Path.GetDirectoryName(DllLocation);
            UserFolder = @"C:\Users\" + Environment.UserName;

            var CutHolesBtnData = new PushButtonData("CutHolesBtnData", "Вырезать", DllLocation, "CutHoles.Command")
            {
                ToolTipImage = new BitmapImage(new Uri(Path.GetDirectoryName(DllLocation) + "\\M1pluginsResource\\cutholes.png", UriKind.Absolute)),
                ToolTip = "Вырезание отверстий в стенах и перекрытиях по заглушкам"
            };
            var CutHolesBtn = Panel.AddItem(CutHolesBtnData) as PushButton;
            CutHolesBtn.LargeImage = new BitmapImage(new Uri(Path.GetDirectoryName(DllLocation) + "\\M1pluginsResource\\cutholes.png", UriKind.Absolute));

            //Panel.AddSeparator();

            return Result.Succeeded;
        }
        public Result OnShutdown(UIControlledApplication application)
        {
            return Result.Succeeded;
        }
    }
}
