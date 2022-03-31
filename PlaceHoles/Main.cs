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

namespace PlaceHoles
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

            var AutoHolesBtnData = new PushButtonData("AutoHolesBtnData", "Расставить", DllLocation, "PlaceHoles.Command")
            {
                ToolTipImage = new BitmapImage(new Uri(Path.GetDirectoryName(DllLocation) + "\\M1pluginsResource\\holes.png", UriKind.Absolute)),
                ToolTip = "Автоматическая расстановка отверстий"
            };
            var AutoHolesBtn = Panel.AddItem(AutoHolesBtnData) as PushButton;
            AutoHolesBtn.LargeImage = new BitmapImage(new Uri(Path.GetDirectoryName(DllLocation) + "\\M1pluginsResource\\holes.png", UriKind.Absolute));

            //Panel.AddSeparator();

            return Result.Succeeded;
        }
        public Result OnShutdown(UIControlledApplication application)
        {
            return Result.Succeeded;
        }

        //private System.Windows.Media.ImageSource PngImageSource(string embeddedPath)
        //{
        //    Stream stream = this.GetType().Assembly.GetManifestResourceStream(embeddedPath);
        //    var decoder = new System.Windows.Media.Imaging.PngBitmapDecoder(stream, BitmapCreateOptions.PreservePixelFormat, BitmapCacheOption.Default);

        //    return decoder.Frames[0];
        //}
    }
}
