using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.Revit.UI;
using Autodesk.Revit.DB;
using Autodesk.Revit.Attributes;
using System.IO;
using System.Windows;
using System.Diagnostics;
using Application = Autodesk.Revit.ApplicationServices.Application;
using Autodesk.Revit.UI.Selection;
using System.Reflection;
using Autodesk.Revit.DB.Mechanical;

namespace PlaceHoles
{
    [Transaction(TransactionMode.Manual), Regeneration(RegenerationOption.Manual)]
    class Command : IExternalCommand
    {
        Document doc { get; set; }
        double wallIndent { get; set; }
        double wallGap { get; set; }
        public Result Execute(ExternalCommandData cmdData, ref string message, ElementSet elements)
        {
              
            Application app = cmdData.Application.Application;
            UIDocument uidoc = cmdData.Application.ActiveUIDocument;
            doc = uidoc.Document;

            CommandView view = new CommandView();
            CommandViewModel vm = (CommandViewModel)view.DataContext;
            view.CommandData = cmdData;
            view.Show();

            return Result.Succeeded;
        }
    }
}
