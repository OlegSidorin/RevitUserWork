using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Globalization;
using System.IO;
using System.Text;
using System.Windows;

namespace Events
{
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    [Autodesk.Revit.Attributes.Regeneration(Autodesk.Revit.Attributes.RegenerationOption.Manual)]
    class SetParameters_ModifyElement : IExternalCommand
    {
        public ElementId ElementId { get; set; }
        public string UserName { get; set; }
        private EventHandler<Autodesk.Revit.UI.Events.IdlingEventArgs> EventHandler { get; set; }
        private UIApplication ApplicationUI { get; set; }
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            return Execute(commandData.Application);
        }
        public Result Execute(UIApplication uiApp)
        {
            UIDocument uiDoc = uiApp.ActiveUIDocument;
            Document doc = uiDoc.Document;
            var app = uiApp.Application;
            ApplicationUI = uiApp;
            EventHandler = new EventHandler<Autodesk.Revit.UI.Events.IdlingEventArgs>(IdleUpdate);
            try
            {
                ApplicationUI.Idling += EventHandler;
            }
            catch (Exception ex)
            {
                Main.WriteToFile(Main.DebugLogFile, DateTime.Now.ToString("dd-MM-yyyy HH:mm") + "\tSetParameters_ModifyElement.ApplicationUI.Idling.Add: " + ex.ToString());
            }
            return Result.Succeeded;
        }
        private void IdleUpdate(object sender, Autodesk.Revit.UI.Events.IdlingEventArgs e)
        {

            UIApplication uiApp = sender as UIApplication;
            Document doc = uiApp.ActiveUIDocument.Document;

            try
            {
                //WriteToFile(DebugLogFile, doc.Title);
                Element element = doc.GetElement(ElementId);
                //MessageBox.Show(doc.Title + " " + element.Name);
                DateTime dateTime = DateTime.Now;
                //dateTime = new DateTime(2022, 3, 2);
                if (element != null)
                {
                    //WriteToFile(DebugLogFile, element.Name);
                    using (Transaction t = new Transaction(doc, "Set parameters"))
                    {
                        t.Start();
                        try 
                        { 
                            var prm = element.LookupParameter("M1_ElementModifier");
                            if (prm != null)
                            {
                                if (!prm.IsReadOnly) _ = prm.Set(UserName);
                            }
                        } 
                        catch (Exception ex) { Main.WriteToFile(Main.DebugLogFile, DateTime.Now.ToString("dd-MM-yyyy HH:mm") + "\tSetParameters_ModifyElement: " + ex.ToString()); }
                        try 
                        { 
                            var prm = element.LookupParameter("M1_ElementModifyDateTime");
                            if (prm != null)
                            {
                                if (!prm.IsReadOnly) _ = prm.Set(dateTime.ToString("dd-MM-yyyy HH:mm"));
                            }
                        } 
                        catch (Exception ex) { Main.WriteToFile(Main.DebugLogFile, DateTime.Now.ToString("dd-MM-yyyy HH:mm") + "\tSetParameters_ModifyElement: " + ex.ToString()); }
                        try
                        {
                            int weekNumber = GetWeekNumber(dateTime);
                            var prm = element.LookupParameter("M1_ElementModifyWeek");
                            if (prm != null)
                            {
                                if (!prm.IsReadOnly) _ = prm.Set(weekNumber);
                            }
                        }
                        catch (Exception ex)
                        {
                            Main.WriteToFile(Main.DebugLogFile, DateTime.Now.ToString("dd-MM-yyyy HH:mm") + "\tSetParameters_ModifyElement: " + ex.ToString());
                        }
                        try
                        {
                            int monthNumber = GetMonthNumber(dateTime);
                            var prm = element.LookupParameter("M1_ElementModifyMonth");
                            if (prm != null)
                            {
                                if (!prm.IsReadOnly) _ = prm.Set(monthNumber);
                            }
                        }
                        catch (Exception ex)
                        {
                            Main.WriteToFile(Main.DebugLogFile, DateTime.Now.ToString("dd-MM-yyyy HH:mm") + "\tSetParameters_ModifyElement: " + ex.ToString());
                        }
                        t.Commit();
                    }
                }

            }
            catch (Exception ex)
            {
                Main.WriteToFile(Main.DebugLogFile, DateTime.Now.ToString("dd-MM-yyyy HH:mm") + "\tSetParameters_ModifyElement: " + ex.ToString());
            }
            try
            {
                ApplicationUI.Idling -= EventHandler;
            }
            catch (Exception ex)
            {
                Main.WriteToFile(Main.DebugLogFile, DateTime.Now.ToString("dd-MM-yyyy HH:mm") + "\tSetParameters_ModifyElement.ApplicationUI.Idling.Remove: " + ex.ToString());
            }
        }

        // This presumes that weeks start with Monday.
        // Week 1 is the 1st week of the year with a Thursday in it.
        private int GetWeekNumber(DateTime time)
        {
            // Seriously cheat.  If its Monday, Tuesday or Wednesday, then it'll 
            // be the same week# as whatever Thursday, Friday or Saturday are,
            // and we always get those right
            DayOfWeek day = CultureInfo.InvariantCulture.Calendar.GetDayOfWeek(time);
            if (day >= DayOfWeek.Monday && day <= DayOfWeek.Wednesday)
            {
                time = time.AddDays(3);
            }

            // Return the week of our adjusted day
            return CultureInfo.InvariantCulture.Calendar.GetWeekOfYear(time, CalendarWeekRule.FirstFourDayWeek, DayOfWeek.Monday);
        }

        private int GetMonthNumber(DateTime time)
        {
            return time.Month;
        }
    }
}
