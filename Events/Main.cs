using System;
using System.IO;
using System.Reflection;
using System.Text;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System.Windows;
using Autodesk.Revit.DB.Events;
using ZetaLongPaths;
using ZLPN = ZetaLongPaths.Native;

namespace Events
{
    public class Main : IExternalApplication
    {
        public static string UserName { get; set; }
        public static string FOPPath { get; } = @"\\ukkalita.local\iptg\Строительно-девелоперский дивизион\М1 Проект\Проекты\10. Отдел информационного моделирования\01. REVIT\00. ФОП\ФОП2019.txt";
        public static string LogFolder { get; } = @"\\ukkalita.local\iptg\Строительно-девелоперский дивизион\М1 Проект\Проекты\3. Обмен\BIM\Корзинка";
        public static string LogFile { get; set; }
        public static string DebugLogFile { get; set; }

        public static string[] ParameterNames { get; set; } = new string[8]
            {
                "M1_ElementModifyDateTime", "M1_ElementCtreateDateTime", "M1_ElementModifier", "M1_ElementCreator", "M1_ElementModifyWeek", "M1_ElementCreateWeek", "M1_ElementModifyMonth", "M1_ElementCreateMonth"
            };

        public Result OnStartup(UIControlledApplication application)
        {
            UserName = Environment.UserName;

            LogFile = LogFolder + $@"\{UserName}_{DateTime.Now:MM-dd}_log.txt";
            DebugLogFile = LogFolder + $@"\{UserName}_{DateTime.Now:MM-dd}_debug.txt";

            if (!ZlpIOHelper.FileExists(LogFile))
            {
                try
                {
                    using (FileStream fs = new FileStream(
                            ZlpIOHelper.CreateFileHandle(
                                LogFile, 
                                ZLPN.CreationDisposition.CreateAlways, 
                                ZLPN.FileAccess.GenericRead | ZLPN.FileAccess.GenericWrite, 
                                ZLPN.FileShare.Read | ZLPN.FileShare.Write), 
                            FileAccess.ReadWrite))
                    {
                        byte[] info = new UTF8Encoding(true).GetBytes("");
                        fs.Write(info, 0, info.Length);
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.ToString());
                }
            }
            if (!File.Exists(DebugLogFile))
            {
                try
                {
                    using (FileStream fs = new FileStream(
                            ZlpIOHelper.CreateFileHandle(
                                DebugLogFile, 
                                ZLPN.CreationDisposition.CreateAlways,
                                ZLPN.FileAccess.GenericRead | ZLPN.FileAccess.GenericWrite,
                                ZLPN.FileShare.Read | ZLPN.FileShare.Write),
                            FileAccess.ReadWrite))
                    {
                        byte[] info = new UTF8Encoding(true).GetBytes("");
                        fs.Write(info, 0, info.Length);
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.ToString());
                }
            }

            try
            {
                application.ControlledApplication.DocumentSynchronizedWithCentral += AppEvent_DocumentSynchronizedWithCentral_Handler;
                application.ControlledApplication.DocumentOpened += AppEvent_DocumentOpened_Handler;
                application.ControlledApplication.DocumentCreated += AppEvent_DocumentCreated_Handler;
                application.ControlledApplication.DocumentSaved += AppEvent_DocumentSaved_Handler;
                application.ControlledApplication.DocumentClosed += AppEvent_DocumentClosed_Handler;
                application.ControlledApplication.FamilyLoadedIntoDocument += AppEvent_FamilyLoadedIntoDocument_Handler;
                application.ControlledApplication.DocumentChanged += AppEvent_DocumentChanged_Handler;
            }
            catch (Exception ex)
            {
                //MessageBox.Show(ex.ToString());
                WriteToFile(DebugLogFile, String.Format("{0}\t{1}", DateTime.Now.ToString("dd-MM-yyyy HH:mm"), ex.ToString()));
            }

            return Result.Succeeded;
        }

        public Result OnShutdown(UIControlledApplication application)
        {
            try
            {
                application.ControlledApplication.DocumentSynchronizedWithCentral -= AppEvent_DocumentSynchronizedWithCentral_Handler;
                application.ControlledApplication.DocumentOpened -= AppEvent_DocumentOpened_Handler;
                application.ControlledApplication.DocumentCreated -= AppEvent_DocumentCreated_Handler;
                application.ControlledApplication.DocumentSaved -= AppEvent_DocumentSaved_Handler;
                application.ControlledApplication.DocumentClosed -= AppEvent_DocumentClosed_Handler;
                application.ControlledApplication.FamilyLoadedIntoDocument -= AppEvent_FamilyLoadedIntoDocument_Handler;
                application.ControlledApplication.DocumentChanged -= AppEvent_DocumentChanged_Handler;
            }
            catch (Exception ex)
            {
                //MessageBox.Show(ex.ToString());
                WriteToFile(DebugLogFile, String.Format("{0}\t{1}", DateTime.Now.ToString("dd-MM-yyyy HH:mm"), ex.ToString()));
            }

            return Result.Succeeded;
        }

        /// <summary>
        /// Обработчик события изменений в проекте Ревит
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private static void AppEvent_DocumentChanged_Handler(object sender, DocumentChangedEventArgs e)
        {
            LogFile = LogFolder + $@"\{UserName}_{DateTime.Now:MM-dd}_log.txt";
            DebugLogFile = LogFolder + $@"\{UserName}_{DateTime.Now:MM-dd}_debug.txt";

            string added = "";

            string modified = "";

            string deleted = "";

            foreach (ElementId eId in e.GetAddedElementIds())
            {
                added += eId.IntegerValue.ToString() + ", ";

                SetParameters_AddElement setParameters = new SetParameters_AddElement();
                setParameters.ElementId = eId;
                setParameters.UserName = e.GetDocument().Application.Username;
                if (sender is UIApplication)
                {
                    try
                    {
                        setParameters.Execute(sender as UIApplication);
                    }
                    catch (Exception ex)
                    {
                        //MessageBox.Show(ex.ToString());
                        WriteToFile(DebugLogFile, String.Format("{0}\t{1}", DateTime.Now.ToString("dd-MM-yyyy HH:mm"), ex.ToString()));
                    }
                }
                else
                {
                    try
                    {
                        setParameters.Execute(new UIApplication(sender as Autodesk.Revit.ApplicationServices.Application));
                    }
                    catch (Exception ex)
                    {
                        //MessageBox.Show(ex.ToString());
                        WriteToFile(DebugLogFile, String.Format("{0}\t{1}", DateTime.Now.ToString("dd-MM-yyyy HH:mm"), ex.ToString()));
                    }
                }

            }

            if (added.Length > 3)
            {
                added = added.Substring(0, added.Length - 2);

                string message = String.Format("{0}\t{1}\tadded:\t{2}", DateTime.Now.ToString("dd-MM-yyyy HH:mm"), UserName, added);

                try
                {
                    WriteToFile(LogFile, message);
                }
                catch (Exception ex)
                {
                    //MessageBox.Show(ex.ToString());
                    WriteToFile(DebugLogFile, String.Format("{0}\t{1}", DateTime.Now.ToString("dd-MM-yyyy HH:mm"), ex.ToString()));
                }

            }

            foreach (var item in e.GetDeletedElementIds())
            {
                deleted += item.IntegerValue.ToString() + ", ";
            }

            if (deleted.Length > 3)
            {
                deleted = deleted.Substring(0, deleted.Length - 2);

                string message = String.Format("{0}\t{1}\tdeleted:\t{2}", DateTime.Now.ToString("dd-MM-yyyy HH:mm"), UserName, deleted);

                try
                {
                    WriteToFile(LogFile, message);
                }
                catch (Exception ex)
                {
                    //MessageBox.Show(ex.ToString());
                    WriteToFile(DebugLogFile, String.Format("{0}\t{1}", DateTime.Now.ToString("dd-MM-yyyy HH:mm"), ex.ToString()));
                }

            }

            foreach (var eId in e.GetModifiedElementIds())
            {
                modified += eId.IntegerValue.ToString() + ", ";

                SetParameters_ModifyElement setParameters = new SetParameters_ModifyElement();
                setParameters.ElementId = eId;
                setParameters.UserName = UserName;
                if (sender is UIApplication)
                {
                    try
                    {
                        setParameters.Execute(sender as UIApplication);
                    }
                    catch (Exception ex)
                    {
                        //MessageBox.Show(ex.ToString());
                        WriteToFile(DebugLogFile, String.Format("{0}\t{1}", DateTime.Now.ToString("dd-MM-yyyy HH:mm"), ex.ToString()));
                    }
                }
                else
                {
                    try
                    {
                        setParameters.Execute(new UIApplication(sender as Autodesk.Revit.ApplicationServices.Application));
                    }
                    catch (Exception ex)
                    {
                        //MessageBox.Show(ex.ToString());
                        WriteToFile(DebugLogFile, String.Format("{0}\t{1}", DateTime.Now.ToString("dd-MM-yyyy HH:mm"), ex.ToString()));
                    }
                }

            }

            if (modified.Length > 3)
            {
                modified = modified.Substring(0, modified.Length - 2);

                string message = String.Format("{0}\t{1}\tmodified:\t{2}", DateTime.Now.ToString("dd-MM-yyyy HH:mm"), UserName, modified);

                try
                {
                    WriteToFile(LogFile, message);
                }
                catch (Exception ex)
                {
                    //MessageBox.Show(ex.ToString());
                    WriteToFile(DebugLogFile, String.Format("{0}\t{1}", DateTime.Now.ToString("dd-MM-yyyy HH:mm"), ex.ToString()));
                }

            }

        }

        /// <summary>
        /// Обработчик события синхронизации с файлом хранилищем
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        public static void AppEvent_DocumentSynchronizedWithCentral_Handler(Object sender, DocumentSynchronizedWithCentralEventArgs e)
        {
            string cloudModelPath = "";
            string revitServerModelPath = "";
            string output = "";

            if (e.Document.IsModelInCloud)
            {
                try
                {
                    cloudModelPath = ModelPathUtils.ConvertModelPathToUserVisiblePath(e.Document.GetCloudModelPath());
                    output = cloudModelPath;
                }
                catch (Exception ex)
                {
                    //MessageBox.Show("isModelInCloud: " + e.Document.IsModelInCloud.ToString() + " and Error with get cloud" + ex.ToString());
                    WriteToFile(DebugLogFile, String.Format("{0}\t{1}", DateTime.Now.ToString("dd-MM-yyyy HH:mm"), ex.ToString()));
                }
            }
            else
            {
                try
                {
                    revitServerModelPath = ModelPathUtils.ConvertModelPathToUserVisiblePath(e.Document.GetWorksharingCentralModelPath()); // revit server or pathname
                    output = revitServerModelPath;
                }
                catch (Exception ex)
                {
                    //MessageBox.Show("Error with Revit Server path " + ex.ToString());
                    WriteToFile(DebugLogFile, String.Format("{0}\t{1}", DateTime.Now.ToString("dd-MM-yyyy HH:mm"), ex.ToString()));
                }
            }

            string message = String.Format("{0}\t{1}\tsync:\t{2}", DateTime.Now.ToString("dd-MM-yyyy HH:mm"), UserName, output);

            try
            {
                WriteToFile(LogFile, message);
            }
            catch (Exception ex)
            {
                //MessageBox.Show(ex.ToString());
                WriteToFile(DebugLogFile, String.Format("{0}\t{1}", DateTime.Now.ToString("dd-MM-yyyy HH:mm"), ex.ToString()));
            }
        }

        /// <summary>
        /// Обработчик создания проекта в Ревит
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        public static void AppEvent_DocumentCreated_Handler(Object sender, DocumentCreatedEventArgs e)
        {
            Autodesk.Revit.ApplicationServices.Application app = e.Document.Application;

            Autodesk.Revit.DB.Document doc = e.Document;

            app.SharedParametersFilename = FOPPath;

            CategorySet catSet = SetCategoryset(app, doc);

            AddParameters(app, doc, ParameterNames, catSet);

            string message = String.Format("{0}\t{1}\tcreated:\t{2}", DateTime.Now.ToString("dd-MM-yyyy HH:mm"), UserName, e.Document.Title + ".rvt");

            try
            {
                WriteToFile(LogFile, message);
            }
            catch (Exception ex)
            {
                //MessageBox.Show(ex.ToString());
                WriteToFile(DebugLogFile, String.Format("{0}\t{1}", DateTime.Now.ToString("dd-MM-yyyy HH:mm"), ex.ToString()));
            }
        }

        /// <summary>
        /// Обработчик события открытия проекта в Ревит
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        public static void AppEvent_DocumentOpened_Handler(Object sender, DocumentOpenedEventArgs e)
        {
            Autodesk.Revit.ApplicationServices.Application app = e.Document.Application;

            Autodesk.Revit.DB.Document doc = e.Document;

            app.SharedParametersFilename = FOPPath;

            CategorySet catSet = SetCategoryset(app, doc);

            AddParameters(app, doc, ParameterNames, catSet);

            string message = String.Format("{0}\t{1}\topened:\t{2}", DateTime.Now.ToString("dd-MM-yyyy HH:mm"), UserName, e.Document.PathName);

            try
            {
                WriteToFile(LogFile, message);
            }
            catch (Exception ex)
            {
                //MessageBox.Show(e.ToString());
                WriteToFile(DebugLogFile, String.Format("{0}\t{1}", DateTime.Now.ToString("dd-MM-yyyy HH:mm"), ex.ToString()));
            }
        }

        /// <summary>
        /// Обработчик события сохранения файла Ревит
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private static void AppEvent_DocumentSaved_Handler(object sender, DocumentSavedEventArgs e)
        {

            string message = String.Format("{0}\t{1}\tsaved:\t{2}", DateTime.Now.ToString("dd-MM-yyyy HH:mm"), UserName, e.Document.PathName);

            try
            {
                WriteToFile(LogFile, message);
            }
            catch (Exception ex)
            {
                //MessageBox.Show(ex.ToString());
                WriteToFile(DebugLogFile, String.Format("{0}\t{1}", DateTime.Now.ToString("dd-MM-yyyy HH:mm"), ex.ToString()));
            }
        }

        /// <summary>
        /// Обработчик события закрытия файла
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private static void AppEvent_DocumentClosed_Handler(object sender, DocumentClosedEventArgs e)
        {

            string message = String.Format("{0}\t{1}\tclosed:\t{2}", DateTime.Now.ToString("dd-MM-yyyy HH:mm"), UserName, e.DocumentId);

            try
            {
                WriteToFile(LogFile, message);
            }
            catch (Exception ex)
            {
                //MessageBox.Show(ex.ToString());
                WriteToFile(DebugLogFile, String.Format("{0}\t{1}", DateTime.Now.ToString("dd-MM-yyyy HH:mm"), ex.ToString()));
            }
        }

        /// <summary>
        /// Обработчик события загрузки семейства в проект Ревит
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        public static void AppEvent_FamilyLoadedIntoDocument_Handler(Object sender, FamilyLoadedIntoDocumentEventArgs e)
        {
            string message = String.Format("{0}\t{1}\tinto:\t{2}\tloaded:\t{3}\tfrom:\t{4}", DateTime.Now.ToString("dd-MM-yyyy HH:mm"), e.Document.Application.Username, e.Document.PathName, e.FamilyName + ".rfa", e.FamilyPath);

            try
            {
                WriteToFile(LogFile, message);
            }
            catch (Exception ex)
            {
                //MessageBox.Show(ex.ToString());
                WriteToFile(DebugLogFile, String.Format("{0}\t{1}", DateTime.Now.ToString("dd-MM-yyyy HH:mm"), ex.ToString()));
            }

        }

        /// <summary>
        /// Метод записи в файл
        /// </summary>
        /// <param name="fileName"></param>
        /// <param name="txt"></param>
        public static void WriteToFile(string fileName, string txt)
        {

            if (!ZlpIOHelper.FileExists(fileName))
            {
                try
                {
                    using (FileStream fs = new System.IO.FileStream(
                                ZlpIOHelper.CreateFileHandle(
                                    fileName, ZLPN.CreationDisposition.CreateAlways,
                                    ZLPN.FileAccess.GenericRead | ZLPN.FileAccess.GenericWrite,
                                    ZLPN.FileShare.Read | ZLPN.FileShare.Write),
                                FileAccess.ReadWrite))
                    {
                        byte[] info = new UTF8Encoding(true).GetBytes("");
                        fs.Write(info, 0, info.Length);
                    }
                }
                catch (Exception ex)
                {
                    //MessageBox.Show(e.ToString());
                }
            } 

            using (var fs = new FileStream(
                    ZlpIOHelper.CreateFileHandle(
                        fileName,
                        ZLPN.CreationDisposition.OpenExisting,
                        ZLPN.FileAccess.GenericWrite,
                        ZLPN.FileShare.None),
                    FileAccess.Write))
            using (var streamWriter = new StreamWriter(fs, new UTF8Encoding(false, true)))
            {
                try
                {
                    fs.Seek(0, SeekOrigin.End);
                    streamWriter.WriteLine(txt);
                }
                catch
                {
                    //MessageBox.Show(e.ToString());
                }
            }
        }

        /// <summary>
        /// Возвращает список категорий элементов для общих параметров
        /// </summary>
        /// <param name="app"></param>
        /// <param name="doc"></param>
        /// <returns></returns>
        public static CategorySet SetCategoryset(Autodesk.Revit.ApplicationServices.Application app, Autodesk.Revit.DB.Document doc)
        {
            CategorySet catSet = app.Create.NewCategorySet();
            Category category;

            // arch
            category = doc.Settings.Categories.get_Item(BuiltInCategory.OST_Walls); catSet.Insert(category);
            category = doc.Settings.Categories.get_Item(BuiltInCategory.OST_Floors); catSet.Insert(category);
            category = doc.Settings.Categories.get_Item(BuiltInCategory.OST_Doors); catSet.Insert(category);
            category = doc.Settings.Categories.get_Item(BuiltInCategory.OST_Windows); catSet.Insert(category);
            category = doc.Settings.Categories.get_Item(BuiltInCategory.OST_Roofs); catSet.Insert(category);
            category = doc.Settings.Categories.get_Item(BuiltInCategory.OST_Ceilings); catSet.Insert(category);
            category = doc.Settings.Categories.get_Item(BuiltInCategory.OST_Columns); catSet.Insert(category);
            category = doc.Settings.Categories.get_Item(BuiltInCategory.OST_Rooms); catSet.Insert(category);

            // struct
            category = doc.Settings.Categories.get_Item(BuiltInCategory.OST_Rebar); catSet.Insert(category);
            category = doc.Settings.Categories.get_Item(BuiltInCategory.OST_AreaRein); catSet.Insert(category);
            category = doc.Settings.Categories.get_Item(BuiltInCategory.OST_StructuralFoundation); catSet.Insert(category);
            category = doc.Settings.Categories.get_Item(BuiltInCategory.OST_StructuralFraming); catSet.Insert(category);
            category = doc.Settings.Categories.get_Item(BuiltInCategory.OST_StructuralColumns); catSet.Insert(category);
            category = doc.Settings.Categories.get_Item(BuiltInCategory.OST_Stairs); catSet.Insert(category);
            category = doc.Settings.Categories.get_Item(BuiltInCategory.OST_StairsRailing); catSet.Insert(category);

            // topo
            category = doc.Settings.Categories.get_Item(BuiltInCategory.OST_Topography); catSet.Insert(category);
            category = doc.Settings.Categories.get_Item(BuiltInCategory.OST_Roads); catSet.Insert(category);

            // generic
            category = doc.Settings.Categories.get_Item(BuiltInCategory.OST_GenericModel); catSet.Insert(category);

            // devices
            category = doc.Settings.Categories.get_Item(BuiltInCategory.OST_FireAlarmDevices); catSet.Insert(category);
            category = doc.Settings.Categories.get_Item(BuiltInCategory.OST_LightingDevices); catSet.Insert(category);
            category = doc.Settings.Categories.get_Item(BuiltInCategory.OST_CommunicationDevices); catSet.Insert(category);
            category = doc.Settings.Categories.get_Item(BuiltInCategory.OST_DataDevices); catSet.Insert(category);
            category = doc.Settings.Categories.get_Item(BuiltInCategory.OST_SecurityDevices); catSet.Insert(category);
            category = doc.Settings.Categories.get_Item(BuiltInCategory.OST_TelephoneDevices); catSet.Insert(category);
            category = doc.Settings.Categories.get_Item(BuiltInCategory.OST_NurseCallDevices); catSet.Insert(category);

            // anno
            category = doc.Settings.Categories.get_Item(BuiltInCategory.OST_Views); catSet.Insert(category);
            category = doc.Settings.Categories.get_Item(BuiltInCategory.OST_Sheets); catSet.Insert(category);

            // electro
            category = doc.Settings.Categories.get_Item(BuiltInCategory.OST_ElectricalFixtures); catSet.Insert(category);
            category = doc.Settings.Categories.get_Item(BuiltInCategory.OST_LightingFixtures); catSet.Insert(category);
            category = doc.Settings.Categories.get_Item(BuiltInCategory.OST_ElectricalCircuit); catSet.Insert(category);
            category = doc.Settings.Categories.get_Item(BuiltInCategory.OST_Conduit); catSet.Insert(category);
            category = doc.Settings.Categories.get_Item(BuiltInCategory.OST_ConduitFitting); catSet.Insert(category);
            category = doc.Settings.Categories.get_Item(BuiltInCategory.OST_CableTray); catSet.Insert(category);
            category = doc.Settings.Categories.get_Item(BuiltInCategory.OST_CableTrayFitting); catSet.Insert(category);

            // equipment
            category = doc.Settings.Categories.get_Item(BuiltInCategory.OST_MechanicalEquipment); catSet.Insert(category);
            category = doc.Settings.Categories.get_Item(BuiltInCategory.OST_ElectricalEquipment); catSet.Insert(category);
            category = doc.Settings.Categories.get_Item(BuiltInCategory.OST_SpecialityEquipment); catSet.Insert(category);
            category = doc.Settings.Categories.get_Item(BuiltInCategory.OST_ZoneEquipment); catSet.Insert(category);
            category = doc.Settings.Categories.get_Item(BuiltInCategory.OST_Sprinklers); catSet.Insert(category);

            // pipes 
            category = doc.Settings.Categories.get_Item(BuiltInCategory.OST_PipeCurves); catSet.Insert(category);
            category = doc.Settings.Categories.get_Item(BuiltInCategory.OST_PipeFitting); catSet.Insert(category);
            category = doc.Settings.Categories.get_Item(BuiltInCategory.OST_PipeAccessory); catSet.Insert(category);
            category = doc.Settings.Categories.get_Item(BuiltInCategory.OST_FlexPipeCurves); catSet.Insert(category);
            category = doc.Settings.Categories.get_Item(BuiltInCategory.OST_PlaceHolderPipes); catSet.Insert(category);
            category = doc.Settings.Categories.get_Item(BuiltInCategory.OST_PlumbingFixtures); catSet.Insert(category);

            // ducts
            category = doc.Settings.Categories.get_Item(BuiltInCategory.OST_DuctCurves); catSet.Insert(category);
            category = doc.Settings.Categories.get_Item(BuiltInCategory.OST_DuctFitting); catSet.Insert(category);
            category = doc.Settings.Categories.get_Item(BuiltInCategory.OST_DuctAccessory); catSet.Insert(category);
            category = doc.Settings.Categories.get_Item(BuiltInCategory.OST_FlexDuctCurves); catSet.Insert(category);
            category = doc.Settings.Categories.get_Item(BuiltInCategory.OST_PlaceHolderDucts); catSet.Insert(category);
            category = doc.Settings.Categories.get_Item(BuiltInCategory.OST_DuctTerminal); catSet.Insert(category);

            return catSet;
        }

        /// <summary>
        /// Метод добавляет 8 общих параметров из ФОП в документ Ревит 
        /// </summary>
        /// <param name="app"></param>
        /// <param name="doc"></param>
        public static void AddParameters(Autodesk.Revit.ApplicationServices.Application app, Document doc, string[] pNames, CategorySet catSet)
        {
            try
            {
                using (Transaction t = new Transaction(doc))
                {
                    t.Start("Add Shared Parameters");
                    foreach (string name in pNames)
                    {
                        DefinitionFile sharedParameterFile = app.OpenSharedParameterFile();
                        DefinitionGroup sharedParameterGroup = sharedParameterFile.Groups.get_Item("14 М1 Общие");
                        Definition sharedParameterDefinition = sharedParameterGroup.Definitions.get_Item(name);
                        ExternalDefinition externalDefinition =
                            sharedParameterGroup.Definitions.get_Item(name) as ExternalDefinition;
                        Guid guid = externalDefinition.GUID;
                        InstanceBinding newIB = app.Create.NewInstanceBinding(catSet);
                        doc.ParameterBindings.Insert(externalDefinition, newIB, BuiltInParameterGroup.PG_IDENTITY_DATA);
                        //SharedParameterElement sp = SharedParameterElement.Lookup(doc, guid);
                        // InternalDefinition def = sp.GetDefinition();
                        // def.SetAllowVaryBetweenGroups(doc, true);
                    }
                    t.Commit();
                }

            }
            catch (Exception ex)
            {
                //MessageBox.Show(ex.ToString());
                WriteToFile(DebugLogFile, String.Format("{0}\t{1}", DateTime.Now.ToString("dd-MM-yyyy HH:mm"), ex.ToString()));
            }
        }

    }
}
