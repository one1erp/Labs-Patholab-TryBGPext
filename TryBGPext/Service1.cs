using Microsoft.Win32;
using Oracle.ManagedDataAccess.Client;
//using Oracle.ManagedDataAccess.Client;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;
using TryBGPext;
using TryBGPextHelpers;

namespace WindowsService1
{
    public partial class Service1// : ServiceBase
    {


        static OracleConnection cnnNaut;
        static string NautConStr;
        static string strBackground;
        static string strBiobankLoginName;
        static string strProgId;
        private static string InputPath;
        private static string OuputPath;

        static void Main(string[] args)
        {

            SetAppSettings();
            CreateConnection();
            READ_XML();
            return;
            OnStart(new string[4]);
        }

        private static void SetAppSettings()
        {
            NautConStr = ConfigurationManager.ConnectionStrings["NautConnectionString"].ConnectionString;
            InputPath = ConfigurationManager.AppSettings["InputPath"];
            OuputPath = ConfigurationManager.AppSettings["OuputPath"];


        }

        private static bool CreateConnection()
        {
            if (string.IsNullOrEmpty(NautConStr))
            {
                Exit("Connection string is NULL");
                return false;
            }
            else
            {
                cnnNaut = new OracleConnection(NautConStr);
                cnnNaut.Open();

                return cnnNaut.State == ConnectionState.Open;
            }
        }

        private static void Exit(string p)
        {
            Console.WriteLine("Exit Program");
            Console.WriteLine(p);
            Console.WriteLine("Press any key to continue");
            Console.Read();
            //CloseDBConnection();
        }

        private static void READ_XML()
        {
            string xmlDir = InputPath;
            var files = Directory.GetFiles(xmlDir, "*.xml");
            foreach (var item in files)
            {
                try
                {
                    string xml = File.ReadAllText(item);
                    var catalog1 = xml.ParseXML<MAIN>();
                    InsertContainerMSG(catalog1);
                    MoveFile(OuputPath, item);
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Error on Parse XML path " + item);
                    Console.WriteLine(ex.Message);
                    Console.WriteLine("Press any key to continue");
                    Console.Read();
                }

            }



        }

        private static void MoveFile(string xmlDir, string file)
        {

            string NameWithoutExtension = Path.GetFileNameWithoutExtension(file);

            FileInfo f = new FileInfo(file);
            var bb = GetCreateMyFolder(xmlDir);
            File.Move(file, Path.Combine(bb.FullName, NameWithoutExtension + DateTime.Now.ToString("yyyy-MM-dd-HH-mm-ss") + ".xml"));
        }


        private static void InsertContainerMSG(MAIN msg)
        {
            var transaction = cnnNaut.BeginTransaction();
            long NewIdNbr;
            try
            {

                string newidSql = "select lims.sq_U_Container_MSG.nextval from dual";
                OracleCommand cmd = new OracleCommand(newidSql, cnnNaut);
                var newIdRes = cmd.ExecuteScalar();
                Console.WriteLine(newidSql);
                var FixDate = GetDate(msg);

                if (newIdRes != null && long.TryParse(newIdRes.ToString(), out NewIdNbr))
                {
                    string insert1 = string.Format("Insert into LIMS.U_CONTAINER_MSG (U_CONTAINER_MSG_ID,NAME,DESCRIPTION,VERSION,VERSION_STATUS) values ({0},'{1}','222','1','A')", NewIdNbr, NewIdNbr);
                    cmd.CommandText = insert1;
                    Console.WriteLine(insert1);
                    cmd.ExecuteNonQuery();
                    string insert2 = string.Format("Insert into LIMS.U_CONTAINER_MSG_USER" +
                        "(U_CONTAINER_MSG_ID,U_DATE,U_CONTAINER_NBR,U_STATUS,U_SENDER,U_MSG_NAME) values ({0},{1},{2},{3},{4},'{5}')"
                        , NewIdNbr, FixDate, msg.TRCONTNUM, msg.TRCONTSTS, msg.TRHOSCODE, msg.XMLNAME);
                    cmd.CommandText = insert2;
                    Console.WriteLine(insert2);

                    cmd.ExecuteNonQuery();
                    foreach (var item in msg.TRREQNUM)
                    {
                        InsertContainerMSG_Entry(item, NewIdNbr);
                    }
                }
                transaction.Commit();
            }
            catch (Exception ex)
            {

                transaction.Rollback();
                Console.WriteLine(ex.Message);
                Console.WriteLine("Press any key to continue");
                Console.Read();
            }



        }

        private static string GetDate(MAIN msg)
        {
            var d = msg.XMLDATE;
            var hr = msg.XMLHR;
            return d + hr;
        }

        private static void InsertContainerMSG_Entry(MAINTRANSPORT row, long headerId)
        {


            long NewIdNbr;




            string newidSql = "select lims.sq_U_Container_MSG_row.nextval from dual";
            OracleCommand cmd = new OracleCommand(newidSql, cnnNaut);
            var newIdRes = cmd.ExecuteScalar();
            Console.WriteLine(newidSql);


            if (newIdRes != null && long.TryParse(newIdRes.ToString(), out NewIdNbr))
            {
                string insertU_CONTAINER_MSG_ROW1 = string.Format("Insert into LIMS.U_CONTAINER_MSG_ROW (U_CONTAINER_MSG_ROW_ID,NAME,DESCRIPTION,VERSION,VERSION_STATUS) values ({0},'{1}','222','1','A')", NewIdNbr, NewIdNbr);
                cmd.CommandText = insertU_CONTAINER_MSG_ROW1;
                Console.WriteLine(insertU_CONTAINER_MSG_ROW1);
                cmd.ExecuteNonQuery();
                string insertU_CONTAINER_MSG_ROW2 = string.Format("Insert into LIMS.U_CONTAINER_MSG_ROW_USER (U_CONTAINER_MSG_ROW_ID,U_REQUEST,u_msg_id) values ({0},'{1}',{2})", NewIdNbr, row.TRREQUEST, headerId);
                cmd.CommandText = insertU_CONTAINER_MSG_ROW2;
                Console.WriteLine(insertU_CONTAINER_MSG_ROW2);

                cmd.ExecuteNonQuery();

            }




        }

        public static DirectoryInfo GetCreateMyFolder(string baseFolder)
        {
            var now = DateTime.Now;
            var yearName = now.ToString("yyyy");
            var monthName = now.ToString("MM");
            var dayName = now.ToString("dd-MM-yyyy");

            var folder = Path.Combine(baseFolder, Path.Combine(yearName, monthName));

            return Directory.CreateDirectory(folder);
        }

        static protected void OnStart(string[] args)
        {
            try
            {

                // Get encrypted connectstring from registry
                RegistryKey nKey = Registry.LocalMachine.OpenSubKey(@"Software\Thermo\Nautilus");
                string strVersion = nKey.GetValue("CurrentVersion").ToString();
                RegistryKey BiomarkKey = nKey.OpenSubKey(strVersion + @"\BiobankLogin Import Filewatcher");
                //     connString = BiomarkKey.GetValue("DBConnectstring").ToString();

                //strBiobankLoginName = Properties.Settings.Default.BiobankLoginFWName;

                strBiobankLoginName = "Biobank Login";
                ////Decrypt connectstring
                ////User Id=??;Password=??;Data Source=??
                //EncryptDecrypt myEncryptor = new EncryptDecrypt();
                //string connString = myEncryptor.Base64Decode_Decrypt(encryptedConnString);

                string connString = "User Id=lims_sys;Password=lims_sys;Data Source=naut";

                // Open database connection
                cnnNaut = OpenDBConnection(connString);

                OracleCommand cmdFileWatcher = new OracleCommand("SELECT u_in_folder, u_file_extension, u_prog_id FROM lims_sys.u_biobank_login_fw_user WHERE u_biobank_login_fw_user.u_biobank_login_fw_id = (select u_biobank_login_fw_id from lims_sys.u_biobank_login_fw where name = '" + strBiobankLoginName + "')", cnnNaut);
                OracleDataReader drFileWatcher = cmdFileWatcher.ExecuteReader();
                string strInFolder = "";
                string strFileExtension = "";
                strProgId = "";

                while (drFileWatcher.Read())
                {
                    strInFolder = drFileWatcher["u_in_folder"].ToString();
                    strFileExtension = drFileWatcher["u_file_extension"].ToString();
                    strProgId = drFileWatcher["u_prog_id"].ToString();
                }

                // Get folder to watch
                //string strInFolder = GetConfigSetting(@"System\Extensions\BiobankLogin Import Interface\In Folder");
                //strBackground = GetConfigSetting(@"System\Extensions\Patient Import Interface\Background Workstation");
                if (strInFolder == "")
                {
                    //EventLog.WriteEntry("Watch folder not specified! Please specify folder in Nautilus configuration settings.", EventLogEntryType.Error);
                    //  this.Stop();
                }
                else
                {
                    //     Console.WriteLine("Start watching folder: " + strInFolder + " with ProgId " + strProgId + " and Service Name " + strBiobankLoginName, EventLogEntryType.Information);
                }

                // First handle any existing files
                string[] files;
                files = Directory.GetFiles(strInFolder, strFileExtension);
                foreach (string fileName in files)
                {
                    CreateBackgroundRecord(fileName);
                }

                // Create a new FileSystemWatcher and set its properties. Only watch xml files.
                FileSystemWatcher watcher = new FileSystemWatcher();
                //watcher.InternalBufferSize = 51200;
                watcher.Path = strInFolder;
                watcher.NotifyFilter = NotifyFilters.FileName;
                watcher.Filter = strFileExtension;

                // Add event handlers.
                watcher.Created += new FileSystemEventHandler(OnCreated);
                watcher.Error += new ErrorEventHandler(OnError);

                // Start watching.
                watcher.EnableRaisingEvents = true;
            }
            catch (Exception ex)
            {
                //EventLog.WriteEntry("Unhandled Error: " + ex.Message, EventLogEntryType.Error);
            }
        }

        static void OnStop()
        {
            CloseDBConnection(cnnNaut);
        }

        static void OnCreated(object source, FileSystemEventArgs e)
        {
            CreateBackgroundRecord(e.FullPath);
        }

        static void OnError(object source, ErrorEventArgs e)
        {
            //EventLog.WriteEntry("Error: " + e.GetException().ToString(), EventLogEntryType.Warning);
        }

        static void CreateBackgroundRecord(string fileName)
        {
            ////EventLog.WriteEntry("Creating Nautilus background record for file: " + fileName);

            //string strParams = "<Filename>" + fileName + "</Filename><BiobankLoginName>" + strBiobankLoginName + "</BiobankLoginName>";

            //OracleCommand cmdExistingRec = new OracleCommand("SELECT 1 FROM lims_sys.background WHERE INSTR (parameter, '" + fileName + "') > 0", cnnNaut);
            //OracleDataReader reader = cmdExistingRec.ExecuteReader();
            //if (!reader.HasRows)
            //{
            //    OracleCommand cmdConnectBG = cnnNaut.CreateCommand();
            //    cmdConnectBG.Connection = cnnNaut;
            //    cmdConnectBG.CommandType = CommandType.StoredProcedure;
            //    cmdConnectBG.CommandText = "LIMS.LIMS_ENV.CONNECT_BACKGROUND_SESSION";
            //    OracleParameter parRetVal = new OracleParameter("ReturnValue", OracleType.Int32);
            //    parRetVal.Direction = ParameterDirection.ReturnValue;
            //    OracleParameter parClientVer = new OracleParameter("ClientVersion", OracleType.VarChar);
            //    parClientVer.Direction = ParameterDirection.Input;
            //    parClientVer.Value = "File Interface";
            //    cmdConnectBG.Parameters.Add(parRetVal);
            //    cmdConnectBG.Parameters.Add(parClientVer);
            //    cmdConnectBG.ExecuteNonQuery();
            //    cmdConnectBG.Dispose();

            //    //if (strBackground.ToString() == "")
            //    //{
            //    OracleCommand cmdInsertBGTask = cnnNaut.CreateCommand();
            //    cmdInsertBGTask.Connection = cnnNaut;
            //    cmdInsertBGTask.CommandType = CommandType.StoredProcedure;
            //    cmdInsertBGTask.CommandText = "LIMS.BACKGROUND_EXTENSION.EXECUTEBACKGROUNDEXTENSION";
            //    parRetVal = new OracleParameter("ReturnValue", OracleType.Int32);
            //    parRetVal.Direction = ParameterDirection.ReturnValue;
            //    OracleParameter parProgId = new OracleParameter("ProgId", OracleType.VarChar);
            //    parProgId.Direction = ParameterDirection.Input;
            //    OracleParameter parDesc = new OracleParameter("Description", OracleType.VarChar);
            //    parDesc.Direction = ParameterDirection.Input;
            //    OracleParameter parParams = new OracleParameter("Params", OracleType.VarChar);
            //    parParams.Direction = ParameterDirection.Input;
            //    //EventLog.WriteEntry("ProgId " + strProgId + " and Service Name " + strBiobankLoginName, EventLogEntryType.Information);
            //    parProgId.Value = strProgId;
            //    parDesc.Value = strProgId;
            //    parParams.Value = strParams;
            //    cmdInsertBGTask.Parameters.Add(parRetVal);
            //    cmdInsertBGTask.Parameters.Add(parProgId);
            //    cmdInsertBGTask.Parameters.Add(parDesc);
            //    cmdInsertBGTask.Parameters.Add(parParams);
            //    cmdInsertBGTask.ExecuteNonQuery();
            //    cmdInsertBGTask.Dispose();


            //    OracleCommand cmdDisconnectBG = cnnNaut.CreateCommand();
            //    cmdDisconnectBG.CommandText = "LIMS.LIMS_ENV.DISCONNECT_SESSION";
            //    cmdDisconnectBG.CommandType = CommandType.StoredProcedure;
            //    cmdDisconnectBG.Connection = cnnNaut;
            //    cmdDisconnectBG.ExecuteNonQuery();
            //    cmdDisconnectBG.Dispose();
            //}
            //reader.Dispose();
            //cmdExistingRec.Dispose();


        }

        static OracleConnection OpenDBConnection(string strConn)
        {
            Console.WriteLine("Database connection opening", EventLogEntryType.Information);
            Console.WriteLine();
            OracleConnection cnn = new OracleConnection(strConn);
            cnn.Open();
            Console.WriteLine("Database connection open", EventLogEntryType.Information);
            OracleCommand cmdSetRole = new OracleCommand("SET ROLE LIMS_USER", cnn);
            cmdSetRole.ExecuteNonQuery();
            Console.WriteLine("Role set", EventLogEntryType.Information);

            return cnn;
        }

        static void CloseDBConnection(OracleConnection cnn)
        {
            if (cnn.State != ConnectionState.Closed)
            {
                cnn.Close();
                Console.WriteLine("Database connection closed", EventLogEntryType.Information);
            }
        }


    }
}
