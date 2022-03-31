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
using Autodesk.Revit.DB.Structure;

namespace CutHoles
{
    [Transaction(TransactionMode.Manual), Regeneration(RegenerationOption.Manual)]
    class Command : IExternalCommand
    {
        Document doc { get; set; }

        const string family_name_hole_walls_round = "М1_ОтверстиеКруглое_Стена";
        const string family_name_hole_walls_rectang = "М1_ОтверстиеПрямоугольное_Стена";
        const string family_name_hole_floors_round = "М1_ОтверстиеКруглое_Перекрытие";
        const string family_name_hole_floors_rectang = "М1_ОтверстиеПрямоугольное_Перекрытие";
        

public Result Execute(ExternalCommandData cmdData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = cmdData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;

            List<FamilyInstance> familyInstances_hole_walls_round = new FilteredElementCollector(doc).OfClass(typeof(FamilyInstance)).WhereElementIsNotElementType().Cast<FamilyInstance>()
                .Where(x => x.Symbol.FamilyName.Contains(family_name_hole_walls_round))
                .ToList();
            List<FamilyInstance> familyInstances_hole_walls_rectang = new FilteredElementCollector(doc).OfClass(typeof(FamilyInstance)).WhereElementIsNotElementType().Cast<FamilyInstance>()
                .Where(x => x.Symbol.FamilyName.Contains(family_name_hole_walls_rectang))
                .ToList();
            List<FamilyInstance> familyInstances_hole_floors_round = new FilteredElementCollector(doc).OfClass(typeof(FamilyInstance)).WhereElementIsNotElementType().Cast<FamilyInstance>()
                .Where(x => x.Symbol.FamilyName.Contains(family_name_hole_floors_round))
                .ToList();
            List<FamilyInstance> familyInstances_hole_floors_rectang = new FilteredElementCollector(doc).OfClass(typeof(FamilyInstance)).WhereElementIsNotElementType().Cast<FamilyInstance>()
                .Where(x => x.Symbol.FamilyName.Contains(family_name_hole_floors_rectang))
                .ToList();


            var floors = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Floors).WhereElementIsNotElementType().ToElements().ToList();

            var walls = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Walls).WhereElementIsNotElementType().ToElements().ToList();

#if DEBUG
            MessageBox.Show($"floors: {floors.Count()}, walls: {walls.Count()}");
#endif


            foreach (var floor in floors)
            {
                using (Transaction tr = new Transaction(doc, $"Cut floor {floor.Name}"))
                {
                    tr.Start();
                    foreach (var fi in familyInstances_hole_floors_round)
                    {
                        try
                        {
                            InstanceVoidCutUtils.AddInstanceVoidCut(doc, floor, fi);
                        }
                        catch (Exception ex)
                        {
#if DEBUG
                            MessageBox.Show(ex.Message);
#endif
                        }

                    }
                    tr.Commit();
                }
                using (Transaction tr = new Transaction(doc, $"Cut floor {floor.Name}"))
                {
                    tr.Start();
                    foreach (var fi in familyInstances_hole_floors_rectang)
                    {
                        try
                        {
                            InstanceVoidCutUtils.AddInstanceVoidCut(doc, floor, fi);
                        }
                        catch (Exception ex)
                        {
#if DEBUG
                            MessageBox.Show(ex.Message);
#endif
                        }
                    }
                    tr.Commit();
                }
            }

            foreach (var wall in walls)
            {
                using (Transaction tr = new Transaction(doc, $"Cut wall {wall.Name}"))
                {
                    tr.Start();
                    foreach (var fi in familyInstances_hole_walls_round)
                    {
                        try
                        {
                            InstanceVoidCutUtils.AddInstanceVoidCut(doc, wall, fi);
                        }
                        catch (Exception ex)
                        {
#if DEBUG
                            MessageBox.Show(ex.Message);
#endif
                        }
                    }
                    tr.Commit();
                }
                using (Transaction tr = new Transaction(doc, $"Cut wall {wall.Name}"))
                {
                    tr.Start();
                    foreach (var fi in familyInstances_hole_walls_rectang)
                    {
                        try
                        {
                            InstanceVoidCutUtils.AddInstanceVoidCut(doc, wall, fi);
                        }
                        catch (Exception ex)
                        {
#if DEBUG
                            MessageBox.Show(ex.Message);
#endif
                        }
                    }
                    tr.Commit();
                }

            }

            return Result.Succeeded;
        }
    }
}
