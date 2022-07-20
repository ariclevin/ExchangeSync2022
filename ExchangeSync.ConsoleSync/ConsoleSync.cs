using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Infragistics.Documents.Excel;
using Microsoft.Xrm.Sdk;

namespace ExchangeSync
{
    public class ConsoleSync
    {
        Core SyncCore = new Core();

        public ConsoleSync()
        {
            StartConsoleSync(false);
        }

        public ConsoleSync(bool exchangeSyncOnly)
        {
            StartConsoleSync(exchangeSyncOnly);
        }

        public ConsoleSync(bool exchangeSyncOnly, bool resyncFailed)
        {
            StartConsoleSync(exchangeSyncOnly, resyncFailed);
        }


        public void StartConsoleSync(bool exchangeSyncOnly, bool resyncFailed = false)
        {
            SyncCore.LoadConfigurationSettings();
            ProgramSetting.SourceApp = SourceApplication.Console;

            Trace.Reset();

            bool isLicensed = SyncCore.LoadLicensingInformation();
            if (isLicensed)
            {
                bool hasAutoSyncLicense = License.AutoSync;
                if (hasAutoSyncLicense)
                {
                    string fqMachineName = Network.FQDN.ToLower();
                    string licenseMachineName = License.ComputerName.ToLower();

                    // Only run this on a licensed machine
                    if (fqMachineName == licenseMachineName)
                    {
                        bool isKeyValid = false;
                        try
                        {
                            Security crypt = new Security();
                            isKeyValid = crypt.ValidateKey(License.LicenseKey);
                        }
                        catch (System.Exception ex)
                        {
                            Trace.AddLog(EventLevel.Error, DateTime.Now, "ConsoleSync", "Licensing", "Error Validating License Key", ex.Message);
                            Console.WriteLine("Error Validating License Key");
                        }

                        if (isKeyValid)
                        {
                            try
                            {
                                StartProcess(exchangeSyncOnly, resyncFailed);
                            }
                            catch (System.Exception ex)
                            {
                                Trace.AddLog(EventLevel.Error, DateTime.Now, "ConsoleSync", "General", ex.Message, string.Empty);
                                Console.WriteLine("Error: {0}", ex.Message);
                                ExportLogFile();
                            }
                        }
                        else
                        {
                            // Check if Trial Version
                            if (ValidateProduct())
                            {
                                try
                                {
                                    StartProcess(exchangeSyncOnly, resyncFailed);
                                }
                                catch (System.Exception ex)
                                {
                                    Trace.AddLog(EventLevel.Error, DateTime.Now, "ConsoleSync", "General", ex.Message, string.Empty);
                                    Console.WriteLine("Error: {0}", ex.Message);
                                    ExportLogFile();
                                }

                            }
                            else
                            {
                                // The trial version has expired
                                Trace.AddLog(EventLevel.Error, DateTime.Now, "InitiateProcess", "General", "Expired Trial Period", "");
                                Console.WriteLine("Your trial version has expired. Please request a valid license key.");
                                ExportLogFile();
                            }
                        }
                    }
                    else
                    {
                        // This is an unlicensed machine
                        Trace.AddLog(EventLevel.Error, DateTime.Now, "ConsoleSync", "General", String.Format("Application can only execute on {0}", licenseMachineName), "");
                        Console.WriteLine("Error. Application can only execute on {0}", licenseMachineName);
                        ExportLogFile();
                    }
                } // Has AutoSync License
                else
                {
                    Trace.AddLog(EventLevel.Information, DateTime.Now, "ConsoleSync", "General", String.Format("You do not have permission to run the Console Sync application"), "");
                    Console.WriteLine("You do not have permission to run the Console Sync application");
                    ExportLogFile();
                }
            }
            else
            {
                Trace.AddLog(EventLevel.Information, DateTime.Now, "ConsoleSync", "General", String.Format("You do not have permission to run the application. Please contact vendor to purchase a valid license"), "");
                Console.WriteLine("You do not have permission to run the application. Please contact vendor to purchase a valid license");
                ExportLogFile();
            }
        }


        private bool ValidateProduct()
        {
            // This function validates that the product can be run on this machine
            bool rc = true;

            DateTime today = DateTime.UtcNow;
            long todayTicks = today.Ticks;

            if (License.Evaluation)
            {
                if (License.EvaluationExpired)
                {
                    Trace.AddLog(EventLevel.Error, DateTime.Now, "ValidateProduct", "General", "CRMExchangeSync Trial Version Expired. Please contact vendor to get a licensed version", "");
                    rc = false;
                }
                else
                {
                    string trialVersionMessage = " [Trial Version. Expires " + License.ExpirationDate.ToShortDateString() + "]";
                    Trace.AddLog(EventLevel.Information, DateTime.Now, "ValidateProduct", "General", trialVersionMessage, "");
                    rc = true;

                }
            }
            return rc;

        }

        private string ExportLogFile()
        {
            string rc = "";
            // Check if logs directory exists
            if (!System.IO.Directory.Exists(@"c:\logs"))
            {
                System.IO.Directory.CreateDirectory(@"C:\logs");
            }

            List<LogRow> list = Trace.RetrieveLog();
            if (list.Count > 0)
            {
                ExcelLog excel = new ExcelLog();
                Workbook book = excel.CreateBook();

                foreach (LogRow row in list)
                {
                    excel.AddRow(book, row.Level, row.LogDateTime, row.Category, row.MethodName, row.CommandName, row.Message, row.Details, row.Parameters);
                }
                rc = excel.SaveBook(book, true);
                Console.WriteLine("Log file saved to {0}", rc);
            }
            return rc;
        }

        private bool EstablishConnections()
        {
            bool rc = true;
            StringBuilder sb = new StringBuilder();

            string defaultProfileValue = Network.HostName.ToUpper();
            string defaultProfileName = defaultProfileValue + " (Default)";
            ProgramSetting.ApplicationProfile = new KeyValuePair<string, string>(defaultProfileValue, defaultProfileName);

            SyncCore.LoadConfigurationSettings();

            CRMEventArgs crmArgs;
            if (ConfigSetting.UseConnectionString == true)
                crmArgs = SyncCore.ConnectToCRM(ConfigSetting.CRMConnectionString, ConfigSetting.CRMCredentials.Password);
            else
            {
                if (!string.IsNullOrEmpty(ConfigSetting.ClientId))
                    crmArgs = SyncCore.ConnectToCRM(ConfigSetting.CRMServiceUrl, ConfigSetting.ClientId, ConfigSetting.ClientSecret);
                else
                    crmArgs = SyncCore.ConnectToCRM(ConfigSetting.AuthenticationProvider.ToInt(), ConfigSetting.CRMCredentials, ConfigSetting.CRMServiceUrl, ConfigSetting.OrganizationName, ConfigSetting.RegionName);
            }
                

            if (crmArgs.Result == true)
            {
                EntityCollection solutions = SyncCore.crm.RetrieveSolutions("crmexchangesync");
                if (solutions.Entities.Count > 0)
                {
                    // Solution Exists
                    string versionNumber = solutions.Entities[0].Attributes["version"].ToString();
                    ProgramSetting.CRMSolutionNumber = versionNumber;
                }

                // Connected
                Trace.AddLog(EventLevel.Information, DateTime.Now, "EstablishConnections", "CRM", "Connected to CRM Server " + ConfigSetting.CRMServiceUrl, string.Empty);
                Console.WriteLine("Connected to CRM Server {0}", ConfigSetting.CRMServiceUrl);

                SyncCore.LoadApplicationSettings();
                if (AppSetting.FieldMappingAction.Value == "CRM")
                {
                    ProgramSetting.Map = SyncCore.LoadDefaultColumnSet();

                    List<string> fields = new List<string>();
                    if ((ProgramSetting.Map != null) && (ProgramSetting.Map.Count > 0))
                    {
                        foreach (FieldMap field in ProgramSetting.Map)
                        {
                            fields.Add(field.CRMFieldName);
                        }
                        SyncCore.crm.FieldMappings = fields;
                    }
                }
            }
            else
            {
                // Error Connecting to CRM Environment
                Trace.AddLog(EventLevel.Error, DateTime.Now, "EstablishConnections", "CRM", "Unable to Connect to CRM Server: " + crmArgs.ConnectionReason, "CRM Connection State was returned as " + crmArgs.ConnectionState.ToString());
                Console.WriteLine("Unable to connect to CRM Server {0}. {1}", ConfigSetting.CRMServiceUrl, crmArgs.ConnectionReason);
                rc = false;
            }

            try
            {
                ExchangeConnectEventArgs exchArgs = SyncCore.ConnectExchange();
                if (exchArgs.Result != true)
                {
                    // Failed to Connect to Exchange
                    Trace.AddLog(EventLevel.Error, DateTime.Now, "EstablishConnections", "Exchange", "Unable to Connect to Exchange Server: " + exchArgs.ConnectionReason, "CRM Connection State was returned as " + crmArgs.ConnectionState.ToString());
                    Console.WriteLine("Unable to connect to Exchange Server {0}. {1}", ConfigSetting.ExchangeServerUrl, exchArgs.ConnectionReason);
                    rc = false;
                }
                else
                {
                    // Connected To Exchange
                    Trace.AddLog(EventLevel.Information, DateTime.Now, "EstablishConnections", "Exchange", "Connected to Exchange Server " + ConfigSetting.ExchangeServerUrl, string.Empty);
                    Console.WriteLine("Connected to Exchange Server {0}", ConfigSetting.ExchangeServerUrl);
                    string serverName = SyncCore.exch.GetServerName();

                }
            }
            catch (System.Exception ex)
            {
                rc = false;
                sb.AppendLine("Could not connect to Exchange Server: " + ex.Message);
            }
            return rc;
        }

        private void StartProcess(bool exchangeSyncOnly = false, bool resyncFailed = false)
        {
            bool isConnected = EstablishConnections();

            if (isConnected)
            {
                DateTime startDateTime = DateTime.Now;
                if (!exchangeSyncOnly)
                    ExecuteChangeTrackingOperations();
                
                if (resyncFailed)
                    RunExchangeReSync();
                else
                    RunExchangeSync();

                EndProcess(startDateTime);
            }
            else
            {
                Trace.AddLog(EventLevel.Warning, DateTime.Now, "EstablishConnections", "General", "Application did not run. Connections not established.", string.Empty);
                Console.WriteLine("Application cannot execute. Connections were not established");
            }
        }

        private void EndProcess(DateTime startDateTime)
        {
            string fileName = ExportLogFile();
            DateTime endDateTime = DateTime.Now;
            // Create User Job History Record

            int errorCount = Trace.Errors;

            Trace.ClearLog();
            Trace.TerminateLog();

            SyncCore.crm.CreateUserJob(Guid.Empty, 1, "Console Sync", startDateTime, endDateTime, fileName, errorCount);
        }

        private void ExecuteChangeTrackingOperations()
        {
            Console.WriteLine("Executing Change Tracking Operations");
            Trace.AddLog(EventLevel.Information, DateTime.Now, "ExecuteChangeTrackingOperations", "General", "Executing Change Tracking Operations", "");
            SyncCore.ExecuteChangeTrackingOperations();
        }

        private void RunExchangeSync()
        {
            
            if (Trace.Null)
                Trace.InitializeLog();
            else
            {
                if (!Trace.Empty)
                    Trace.ClearLog();
            }

            Trace.AddLog(EventLevel.Information, DateTime.Now, "RunSynchronization", "General", "Starting Sync Process", "");
            Console.WriteLine("Starting Synchronization Process");

            EntityCollection results = SyncCore.crm.RetrieveSyncLists(ProgramSetting.ApplicationProfile.Value, "");
            
            if (results.Entities.Count > 0)
            {
                SyncCore.UpdateSyncStart();
                foreach (var result in results.Entities)
                {
                    bool isActive = result.Contains("xrm_autosyncstatus") ? (bool)result.Attributes["xrm_autosyncstatus"] : false;

                    if (isActive)
                    {
                        Guid syncListId = result.Id;
                        Guid entityId = new Guid(result.Attributes["xrm_entityid"].ToString());
                        string entityName = result.Attributes["xrm_entityname"].ToString();
                        // row.Cells[3].Value = result.Attributes["xrm_entitydisplayname"].ToString();
                        string listName = result.Attributes["xrm_listname"].ToString();
                        string listTypeName = "Contact";
                        int? listType = ((OptionSetValue)(result.Attributes["xrm_objecttypecode"])).Value;
                        switch (listType)
                        {
                            case 1:
                                listTypeName = "Account";
                                break;
                            case 2:
                                listTypeName = "Contact";
                                break;
                            case 4:
                                listTypeName = "Lead";
                                break;
                            default:
                                listTypeName = "Contact";
                                break;
                        }

                        string exchangeGroupName = result.Attributes["xrm_exchangegroupname"].ToString();
                        DateTime lastSync = result.Contains("xrm_lastsync") ? (DateTime)result["xrm_lastsync"] : DateTime.MinValue;
                        try
                        {
                            SyncCore.ExecuteSync(syncListId, entityId, entityName, listName, exchangeGroupName, lastSync, listTypeName);
                        }
                        catch (System.Exception ex)
                        {
                            if (ex.Message == "The underlying connection was closed: A connection that was expected to be kept alive was closed by the server.")
                            {
                                // Reconnect
                                bool isConnected = EstablishConnections();
                                if (!isConnected)
                                {
                                    Trace.AddLog(EventLevel.Critical, DateTime.Now, "Connection", "RunExchangeSync", "Unable to re-establish connections to CRM and Exchange");
                                }
                            }
                        }
                        
                        Console.WriteLine("Finished Executing List {0}", listName);
                    }
                }
                SyncCore.UpdateSyncEnd();
            }
        }

        private void RunExchangeReSync()
        {

            if (Trace.Null)
                Trace.InitializeLog();
            else
            {
                if (!Trace.Empty)
                    Trace.ClearLog();
            }

            Trace.AddLog(EventLevel.Information, DateTime.Now, "RunSynchronization", "General", "Starting Sync Process", "");
            Console.WriteLine("Starting Resynchronization Process");

            if (!string.IsNullOrEmpty(AppSetting.LastRunStart.Value))
            {
                DateTime lastSyncStart = Convert.ToDateTime(AppSetting.LastRunStart.Value);
                Console.WriteLine("Last Sync Start Date: " + lastSyncStart.ToString());
                EntityCollection results = SyncCore.crm.RetrieveSyncLists(ProgramSetting.ApplicationProfile.Value, lastSyncStart);

                if (results.Entities.Count > 0)
                {
                    SyncCore.UpdateSyncStart();
                    foreach (var result in results.Entities)
                    {
                        bool isActive = result.Contains("xrm_autosyncstatus") ? (bool)result.Attributes["xrm_autosyncstatus"] : false;

                        if (isActive)
                        {
                            Guid syncListId = result.Id;
                            Guid entityId = new Guid(result.Attributes["xrm_entityid"].ToString());
                            string entityName = result.Attributes["xrm_entityname"].ToString();
                            // row.Cells[3].Value = result.Attributes["xrm_entitydisplayname"].ToString();
                            string listName = result.Attributes["xrm_listname"].ToString();
                            string listTypeName = "Contact";
                            int? listType = ((OptionSetValue)(result.Attributes["xrm_objecttypecode"])).Value;
                            switch (listType)
                            {
                                case 1:
                                    listTypeName = "Account";
                                    break;
                                case 2:
                                    listTypeName = "Contact";
                                    break;
                                case 4:
                                    listTypeName = "Lead";
                                    break;
                                default:
                                    listTypeName = "Contact";
                                    break;
                            }

                            string exchangeGroupName = result.Attributes["xrm_exchangegroupname"].ToString();
                            DateTime lastSync = result.Contains("xrm_lastsync") ? (DateTime)result["xrm_lastsync"] : DateTime.MinValue;

                            SyncCore.ExecuteSync(syncListId, entityId, entityName, listName, exchangeGroupName, lastSync, listTypeName);
                            Console.WriteLine("Finished Executing List {0}", listName);
                        }
                    }
                    SyncCore.UpdateSyncEnd();
                }

            }
        }


    }
}
