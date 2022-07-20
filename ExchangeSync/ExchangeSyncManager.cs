using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Management.Automation.Runspaces;
using System.Net;
using System.Reflection;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using System.Xml;
using Infragistics.Win.UltraWinToolbars;
using Infragistics.Win.UltraMessageBox;
using Infragistics.Documents.Excel;

using Microsoft.Xrm.Sdk;

namespace ExchangeSync
{
    public partial class ExchangeSyncManager : Form
    {
        public ExchangeSyncManager()
        {
            InitializeComponent();
        }

        public static Core SyncCore { get; set; }

        #region Startup Functions

        private void frmSync_Load(object sender, EventArgs e)
        {
            Trace.InitializeLog();
            Trace.AddLog(EventLevel.Verbose, DateTime.Now, "App", "Form Load", "Entered frmSync_Load");

            SyncCore = new Core();
            ProgramSetting.SourceApp = SourceApplication.Windows;

            Trace.AddLog(EventLevel.Verbose, DateTime.Now, "App", "Form Load", "Loading Licensing Information", "SyncCore.LoadLicensingInformation()");
            SyncCore.LoadLicensingInformation();
            Trace.AddLog(EventLevel.Verbose, DateTime.Now, "App", "Form Load", "Loading Configuration Settings", "SyncCore.LoadConfigurationSettings()");
            SyncCore.LoadConfigurationSettings();

            CultureInfo culture = new CultureInfo(ConfigSetting.Locale);
            Thread.CurrentThread.CurrentCulture = culture;

            if (ProgramSetting.ApplicationProfile.Key.ToUpper() != Network.HostName.ToUpper())
            {
                Trace.AddLog(EventLevel.Verbose, DateTime.Now, "App", "Form Load", "Loading Profile Configuration Settings", "SyncCore.LoadConfigurationSettings(" + ProgramSetting.ApplicationProfile.Key + ")");
                SyncCore.LoadConfigurationSettings(ProgramSetting.ApplicationProfile.Key);
            }

            lblCRMServer.Text = Helper.CRMServiceUrlToServerUrl(ConfigSetting.CRMServiceUrl); 
            lblExchangeServer.Text = ConfigSetting.ExchangeServerUrl != "" ? ConfigSetting.ExchangeServerUrl : "Exchange Online";

            PopupControlContainerTool popupSettings;
            popupSettings = (PopupControlContainerTool)utmSyncToolbar.Tools["AboutSettings"];
            popupSettings.Control = new AccountInfo();

            popupSettings = (PopupControlContainerTool)utmSyncToolbar.Tools["Confirmation"];
            popupSettings.Control = new Confirmation();

            popupSettings = (PopupControlContainerTool)utmSyncToolbar.Tools["toolFeedback"];
            popupSettings.Control = new Feedback();

            popupSettings = (PopupControlContainerTool)utmSyncToolbar.Tools["HistorySettings"];
            popupSettings.Control = new History();

            string fqdnMachine = Network.FQDN.ToLower();
            string licenseMachine = License.ComputerName.ToLower();

            Trace.AddLog(EventLevel.Verbose, DateTime.Now, "App", "Form Load", "Validate License for " + licenseMachine);
            if (fqdnMachine == licenseMachine)
            {
                // Application Valid to be running on this machine
                string licenseKey = License.LicenseKey;
                if (!string.IsNullOrEmpty(licenseKey))
                {
                    this.Text += " [" + License.LicenseType + " Edition]";
                    switch (License.LicenseType.ToLower())
                    {
                        case "evaluation":
                            if (ValidateProduct(popupSettings))
                            {
                                Application.DoEvents();
                                utmSyncToolbar.Enabled = false;
                                activityIndicator.Start();
                                worker.RunWorkerAsync();
                            }
                            else
                            {
                                ShowSettingsOnly();
                            }
                            break;
                        case "community":   
                            if ((ConfigSetting.CRMServiceUrl.Contains("dynamics.com")) && (ConfigSetting.ExchangeServerVersion == ExchangeServerType.Exchange_Online_365))
                            {
                                Application.DoEvents();
                                ShowCommunityEditionButtonsOnly();
                                utmSyncToolbar.Enabled = false;
                                activityIndicator.Start();
                                worker.RunWorkerAsync();
                            }
                            else
                            {
                                string message = "Application can only be run against CRM and Exchange Online in Community Edition";
                                MessageBox.Show(message, "Community Edition Error", MessageBoxButtons.OK);
                                ShowSettingsOnly();
                            }
                            break;
                        case "enterprise":
                            if (!string.IsNullOrEmpty(License.LicenseKey))
                            {
                                Security crypt = new Security();
                                if (crypt.ValidateKey(License.LicenseKey))
                                {
                                    Application.DoEvents();
                                    utmSyncToolbar.Enabled = false;
                                    activityIndicator.Start();
                                    worker.RunWorkerAsync();
                                }
                                else
                                {
                                    if (ValidateProduct(popupSettings))
                                    {
                                        Application.DoEvents();
                                        utmSyncToolbar.Enabled = false;
                                        activityIndicator.Start();
                                        worker.RunWorkerAsync();
                                    }
                                }
                            }
                            break;
                        default:
                            break;
                    }
                }              
            }
            else
            {
                string message = "Application can only be run on " + licenseMachine;
                MessageBox.Show(message, "Invalid Computer Name.", MessageBoxButtons.OK);
                ShowSettingsOnly();
            }
        }

        private void worker_DoWork(object sender, DoWorkEventArgs e)
        {
            EstablishConnections();
        }

        private void worker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            Trace.AddLog(EventLevel.Verbose, DateTime.Now, "App", "Worker Process", "Populating Sync List", "worker_RunWorkerCompleted");
            activityIndicator.Stop();
            PopulateSyncList();
            utmSyncToolbar.Enabled = true;
            activityIndicator.Visible = false;

            SyncCore.ContactsCounted += SyncCore_ContactsCounted;
            SyncCore.ContactSynched += SyncCore_ContactSynched;
        }

        private void SyncCore_ContactsCounted(object sender, ContactCountEventArgs e)
        {
            tspbStatus.Minimum = 0;
            tspbStatus.Maximum = e.TotalContacts;
            tspbStatus.Value = 0;

        }

        private void SyncCore_ContactSynched(object sender, EventArgs e)
        {
            tspbStatus.Value++;
        }

        private bool ValidateProduct(PopupControlContainerTool settings)
        {
            // This function validates that the product can be run on this machine
            bool rc = true;

            DateTime today = DateTime.UtcNow;
            long todayTicks = today.Ticks;

            if (License.Evaluation)
            {

                if (License.EvaluationExpired)
                {
                    this.Text = string.Format("{0} [Trial Version (Expired)]", this.Text);
                    MessageBox.Show("The trial period of your CRM Exchange Sync product has expired. Please Register your product.", "Trial Expired", MessageBoxButtons.OK);
                    ShowExpiredTrial();
                    rc = false;
                }
                else
                {
                    this.Text = string.Format("{0} [Trial Version (Expires on {1})]", this.Text, License.ExpirationDate.ToShortDateString());
                }
            }
            else
            {
                if (License.ExpirationDate < DateTime.Now)
                {
                    MessageBox.Show("The trial period of your CRM Exchange Sync product has expired. Please Register your product.", "Trial Expired", MessageBoxButtons.OK);
                    ShowExpiredTrial();
                    rc = false;
                }
            }
            return rc;
        }
        
        
        public void EstablishConnections()
        {
            bool hasErrors = false;
            StringBuilder sb = new StringBuilder();

            // activityIndicator.Text = "Establishing Connection with CRM Server";

            CRMEventArgs crmArgs;
            if (ConfigSetting.UseConnectionString == true)
            {
                Trace.AddLog(EventLevel.Verbose, DateTime.Now, "App", "Connect", "Using Connection String", "SyncCore.ConnectToCRM");
                crmArgs = SyncCore.ConnectToCRM(ConfigSetting.CRMConnectionString, ConfigSetting.CRMCredentials.Password);
                // lblCRMServer.Text += "*";
            }
            else
            {
                if (ConfigSetting.AuthenticationProvider.ToInt() == 3)
                {
                    Trace.AddLog(EventLevel.Verbose, DateTime.Now, "App", "Connect", "Using Client Credentials", "SyncCore.ConnectToCRM");
                    crmArgs = SyncCore.ConnectToCRM(ConfigSetting.CRMServiceUrl, ConfigSetting.ClientId, ConfigSetting.ClientSecret);
                }
                else
                {
                    Trace.AddLog(EventLevel.Verbose, DateTime.Now, "App", "Connect", "Using Credentials", "SyncCore.ConnectToCRM");
                    crmArgs = SyncCore.ConnectToCRM(ConfigSetting.AuthenticationProvider.ToInt(), ConfigSetting.CRMCredentials, ConfigSetting.CRMServiceUrl, ConfigSetting.OrganizationName, ConfigSetting.RegionName);
                }
            }

            if (crmArgs.Result == true)
            {
                Trace.AddLog(EventLevel.Verbose, DateTime.Now, "App", "Connect", "Validating Solution Version", "SyncCore.crm.RetrieveSolutions");
                EntityCollection solutions = SyncCore.crm.RetrieveSolutions("crmexchangesync");
                if (solutions.Entities.Count > 0)
                {
                    // Solution Exists
                    string versionNumber = solutions.Entities[0].Attributes["version"].ToString();
                    ProgramSetting.CRMSolutionNumber = versionNumber;
                }

                // Connected
                tslConnection.ToolTipText = "Connected to CRM Server " + ConfigSetting.CRMServiceUrl;
                Trace.AddLog(EventLevel.Verbose, DateTime.Now, "App", "Connect", "Loading Application Settings", "SyncCore.LoadApplicationSettings()");
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
                ShowMessage("ConnectionError", "Unable to Connect to CRM Server: " + crmArgs.ConnectionReason, "CRM Connection State was returned as " + crmArgs.ConnectionState.ToString(), MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }

            try
            {
                // activityIndicator.Text = "Establishing Connection with Exchange Server";

                Trace.AddLog(EventLevel.Verbose, DateTime.Now, "App", "Connect", "Connecting to Exchange Server", "SyncCore.ConnectExchange()");
                ExchangeConnectEventArgs exchArgs = SyncCore.ConnectExchange();
                if (exchArgs.Result != true)
                {
                    // Failed to Connect to Exchange
                    Trace.AddLog(EventLevel.Error, DateTime.Now, "App", "Connect", "Connecting to Exchange Server", "Connection to Exchange failed: " + exchArgs.ConnectionReason);
                    ShowMessage("ConnectionError", "Unable to Connect to Exchange Server: " + exchArgs.ConnectionReason, "Exchange Connection State was returned as " + exchArgs.ConnectionState.ToString(), MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                }
                else
                {
                    // Connected To Exchange
                }

                string serverName = SyncCore.exch.GetServerName();
                Trace.AddLog(EventLevel.Verbose, DateTime.Now, "App", "Connect", "Connected to Exchange Server", "Exchange Server: " + serverName);
            }
            catch (System.Exception ex)
            {
                hasErrors = true;
                sb.AppendLine("Could not connect to Exchange Server: " + ex.Message);
                Trace.AddLog(EventLevel.Error, DateTime.Now, "App", "Connect", "Connecting to Exchange Server", "Connection to Exchange failed: " + ex.Message);
            }

            if (!hasErrors)
            {
                // tslConnection.Image = ExchangeSync.Properties.Resources.connect1;
                try
                {
                    // PopulateSyncList();
                }
                catch (System.Exception ex)
                {
                    hasErrors = true;
                    sb.AppendLine("Error Populating Sync List: " + ex.Message);
                    MessageBox.Show("Error Populating Sync List: " + ex.Message);
                }
            }
            else
            {
                // tslConnection.Image = ExchangeSync.Properties.Resources.connect3;
                
                string message = "The following errors were encountered while trying to connect to your application\n\n";
                message += sb.ToString();
                message += "\n\n";
                message += "Please fix your configuration files and try connecting again.\nOnly configuration pages will show.";
                MessageBox.Show(message, "Configuration Errors Found", MessageBoxButtons.OK);
                ShowSettingsOnly();
            }

        }

        private void ShowSettingsOnly()
        {
            int totalTools = utmSyncToolbar.Tools.Count;
            for (int i = 0; i < totalTools; i++)
            {
                switch (utmSyncToolbar.Tools[i].Key)
                {
                    case "AboutSettings":
                    case "RegisterSettings":
                    case "ExchangeSettings":
                    case "CRMSettings":
                        break;
                    default:
                        utmSyncToolbar.Tools[i].SharedProps.Enabled = false;
                        break;
                }
            }

        }

        private void ShowCommunityEditionButtonsOnly()
        {
            int totalTools = utmSyncToolbar.Tools.Count;
            for (int i = 0; i < totalTools; i++)
            {
                switch (utmSyncToolbar.Tools[i].Key)
                {
                    case "toolCompare":
                    case "toolExportExchangeGroups":
                    case "toolDetectDuplicates":
                    case "toolUpdateExchangeAliases":
                    case "toolSetAppSettings":
                    case "toolClearExchangeAliases":
                    case "toolCRMConnection":
                    case "toolExchangeConnection":
                    case "toolDeleteExchangeContacts":
                    case "keyDeleteExchangeContacts":
                        utmSyncToolbar.Tools[i].SharedProps.Enabled = false;
                        break;
                    default:
                        utmSyncToolbar.Tools[i].SharedProps.Enabled = true;
                        break;
                }
            }
        }

        private void ShowExpiredTrial()
        {
            int totalTools = utmSyncToolbar.Tools.Count;
            for (int i = 0; i < totalTools; i++)
            {
                switch (utmSyncToolbar.Tools[i].Key)
                {
                    case "AboutSettings":
                    case "RegisterSettings":
                        break;
                    default:
                        utmSyncToolbar.Tools[i].SharedProps.Enabled = false;
                        break;
                }
            }
            
        }


        private void PopulateSyncList(bool showAll = true)
        {
            if (dgvSyncList.Rows.Count > 0)
            {
                for (int i = dgvSyncList.Rows.Count - 1; i >= 0; i--)
                {
                    dgvSyncList.Rows.RemoveAt(i);
                }
            }

            Trace.AddLog(EventLevel.Verbose, DateTime.Now, "App", "PopulateSyncList", "Calling Populate Sync Lists", "Application Profile: " + ProgramSetting.ApplicationProfile.Value);
            EntityCollection results = SyncCore.crm.RetrieveSyncLists(ProgramSetting.ApplicationProfile.Value, "");
            
            Trace.AddLog(EventLevel.Verbose, DateTime.Now, "App", "PopulateSyncList", "Populate Sync Lists Called", "Total Lists: " + results.Entities.Count.ToString());
            if (results.Entities.Count > 0)
            {
                foreach (var result in results.Entities)
                {
                    bool isActive = result.Contains("xrm_liststatus") ? (bool)result.Attributes["xrm_liststatus"] : false;

                    bool showRow = false;
                    if (showAll)
                        showRow = true;
                    else
                    {
                        if (isActive)
                            showRow = true;
                        else
                            showRow = false;
                    }

                    if (showRow)
                    {
                        int rowid = dgvSyncList.Rows.Add();
                        DataGridViewRow row = (DataGridViewRow)dgvSyncList.Rows[rowid];
                        row.Cells[SyncListColumn.SyncListId.ToInt()].Value = result.Id.ToString();
                        row.Cells[SyncListColumn.ObjectId.ToInt()].Value = result.Attributes["xrm_entityid"].ToString();
                        row.Cells[SyncListColumn.ObjectTypeName.ToInt()].Value = result.Attributes["xrm_entityname"].ToString();
                        row.Cells[SyncListColumn.ObjectType.ToInt()].Value = result.Attributes["xrm_entitydisplayname"].ToString();
                        row.Cells[SyncListColumn.ListName.ToInt()].Value = result.Attributes["xrm_listname"].ToString();
                        int? listType = ((OptionSetValue)(result.Attributes["xrm_objecttypecode"])).Value;
                        switch (listType)
                        {
                            case 1:
                                row.Cells[SyncListColumn.ListType.ToInt()].Value = "Account";
                                break;
                            case 2:
                                row.Cells[SyncListColumn.ListType.ToInt()].Value = "Contact";
                                break;
                            case 4:
                                row.Cells[SyncListColumn.ListType.ToInt()].Value = "Lead";
                                break;
                            default:
                                row.Cells[SyncListColumn.ListType.ToInt()].Value = "Contact";
                                break;
                        }
                        row.Cells[SyncListColumn.ExchangeGroupName.ToInt()].Value = result.Attributes["xrm_exchangegroupname"].ToString();
                        int groupType = result.Contains("xrm_exchangegrouptypecode") ? ((OptionSetValue)(result.Attributes["xrm_exchangegrouptypecode"])).Value : 1;
                        row.Cells[SyncListColumn.GroupType.ToInt()].Value = Enum.GetName(typeof(GroupType), groupType);
                            
                        if (result.Contains("xrm_lastsync"))
                        {
                            string dateTimeValue = ((DateTime)(result.Attributes["xrm_lastsync"])).ToLocalTime().ToShortDateString();
                            dateTimeValue += " " + ((DateTime)(result.Attributes["xrm_lastsync"])).ToLocalTime().ToShortTimeString();
                            row.Cells[SyncListColumn.LastSync.ToInt()].Value = dateTimeValue;
                        }
                        row.Cells[SyncListColumn.SyncStatus.ToInt()].Value = (bool?)result.Attributes["xrm_liststatus"];
                        row.Cells[SyncListColumn.AutoSyncStatus.ToInt()].Value = (bool?)result.Attributes["xrm_autosyncstatus"];
                    }
                }
            }
        }

        #endregion


        private void utmSyncToolbar_ToolClick(object sender, Infragistics.Win.UltraWinToolbars.ToolClickEventArgs e)
        {
            DialogResult dr;
            switch (e.Tool.Key)
            {
                case "RunSync":    // PopupMenuTool
                case "toolRunFullSync":    // ButtonTool
                    // Place code here
                    RunSync(false);
                    break;

                case "toolRunTestSync":    // ButtonTool
                    // Place code here
                    RunSync(true);
                    break;

                case "toolDeleteExchangeContacts":
                case "keyDeleteExchangeContacts":
                    DialogResult drDeleteExchangeContacts = ShowMessage("Delete Exchange Contacts?", "Clicking OK will Delete All the Contacts from the Change Tracking Operations in CRM. If you are not sure about this, please Cancel and review the data in CRM.", "", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning);
                    if (drDeleteExchangeContacts == DialogResult.OK)
                    {
                        Trace.AddLog(EventLevel.Information, DateTime.Now, "ToolClick", "General", "Start Executing Change Tracking Operations", "");
                        SyncCore.ExecuteChangeTrackingOperations();
                        Trace.AddLog(EventLevel.Information, DateTime.Now, "ToolClick", "General", "End Executing Change Tracking Operations", "");
                        tslCurrent.Text = "Completed Change Tracking Operations";
                    }
                    break;

                case "toolCompare": // ButtonTool
                case "toolCompareSelected":
                    CompareList();
                    break;

                case "toolCompareAll":
                    CompareAll();
                    break;

                case "AddNewGroup":    // ButtonTool
                    AddEditGroup(true);
                    break;

                case "toolEditGroup":   // ButtonTool
                    AddEditGroup(false);
                    break;

                case "toolAddMultipleGroups": // Button Tool
                    AddMultiple();
                    break;

                case "RemoveGroup":    // ButtonTool
                    RemoveGroup();
                    break;

                case "ChangeStatus":    // PopupMenuTool
                    // Place code here
                    break;

                case "Enable":    // ButtonTool
                    ChangeStatus(true);
                    break;

                case "Disable":    // ButtonTool
                    ChangeStatus(false);
                    break;

                case "toolEnableAll":
                    ChangeStatusAll(true);
                    break;

                case "toolDisableAll":
                    ChangeStatusAll(false);
                    break;

                case "toolDisplayActive":
                    StateButtonTool tool = (StateButtonTool)e.Tool;
                    ChangeDataGridDisplay(tool.Checked);
                    break;

                case "toolExportLog":
                    if (!Trace.Null)
                    {
                        if (!Trace.Empty)
                            ExportLogFile();
                        else
                            MessageBox.Show("Log File is empty.", "Log File", MessageBoxButtons.OK);
                    }
                    else
                        MessageBox.Show("Log File has not been initialized.", "Log File", MessageBoxButtons.OK);
                    break;

                case "toolViewLog":
                    ViewLog();
                    break;

                case "toolClearLog":
                    if (!Trace.Null)
                        Trace.ClearLog();
                    break;

                case "toolDuplicateDetection":
                    utmSyncToolbar.Ribbon.Tabs["tabDuplicates"].Visible = true;
                    utmSyncToolbar.Ribbon.ContextualTabGroups[0].Visible = true;
                    utmSyncToolbar.Ribbon.SelectedTab = utmSyncToolbar.Ribbon.Tabs["tabDuplicates"];
                    break;

                case "toolCloseDuplicateDetection":
                    StateButtonTool showTabs = (StateButtonTool)utmSyncToolbar.Ribbon.Tabs["Administration"].Groups["groupCRM"].Tools["toolDuplicateDetection"];
                    showTabs.Checked = false;
                    utmSyncToolbar.Ribbon.Tabs["tabDuplicates"].Visible = false;
                    utmSyncToolbar.Ribbon.ContextualTabGroups[0].Visible = false;
                    utmSyncToolbar.Ribbon.SelectedTab = utmSyncToolbar.Ribbon.Tabs["Administration"];
                    // dgvDuplicates.Visible = false;
                    dgvSyncList.Visible = true;
                    // ugSyncLog.Visible = false;
                    break;

                case "toolDetectDuplicates":
                    // dgvDuplicates.Visible = true;
                    // dgvSyncList.Visible = false;
                    // ugSyncLog.Visible = false;
                    DetectDuplicates();
                    break;

                case "toolQuickValidation":
                    ValidateSync(true);
                    break;
                case "toolUpdateExchangeAliases":
                    // UpdateExchangeAliases(false);
                    break;

                case "toolClearExchangeAliases":
                    // UpdateExchangeAliases(true);
                    break;

                case "toolClearAllExchangeAliases":
                    ClearAllExchangeAliases();
                    break;

                case "toolChangeDuplicateUpdateStatus":
                    // ChangeRowStatus();
                    break;

                case "toolCRMConnection":
                    // ConnectToCRMTest();
                    break;

                case "toolExchangeConnection":
                    // ConnectToExchangeTest();
                    break;

                case "toolReconnect":
                    utmSyncToolbar.Enabled = false;
                    activityIndicator.Start();
                    worker.RunWorkerAsync();
                    break;

                case "ExchangeSettings":    // ButtonTool
                    // Place code here
                    break;

                case "CRMSettings":    // ButtonTool
                    // Place code here
                    break;

                case "AboutSettings":    // ButtonTool
                    // Place code here
                    break;
                case "ExportExchangeGroups":
                    ExportExchangeGroups();
                    break;
                case "toolSetAppSettings":
                    dr = MessageBox.Show("Calling this method will Update all and Create missing Application Settings in SyncCore.crm. Are you sure you want to continue?", "Create/Update Application Settings", MessageBoxButtons.YesNo);
                    if (dr == DialogResult.Yes)
                    {
                        // Helper.SetResetApplicationSettings();
                        // Reload from Excel and Update
                    }

                    break;
                case "toolOptions":
                    ShowOptions();
                    break;
                case "toolCloseRunspaces":

                    break;
                case "Exit":
                    dr = MessageBox.Show("Are you sure you want to Quit?", "Quit", MessageBoxButtons.YesNo);
                    if (dr == DialogResult.Yes)
                        Application.Exit();
                    break;
            }

        }

        private void ShowOptions()
        {
            Options frm = new Options();
            frm.ShowDialog(this);
        }

        private void RunSync(bool isTest = false)
        {
            if (Trace.Null)
                Trace.InitializeLog();
            else
            {
                if (!Trace.Empty)
                    Trace.ClearLog();
            }

            // Trace.AddLogHeader();
            Trace.AddLog(EventLevel.Information, DateTime.Now, "RunSync", "General", "Synchronization Process Started", string.Empty);
            // Trace.AddLog("Run Sync Process Started");

            string footer = "";
            MessageBoxIcon icon = MessageBoxIcon.None;
            if (!isTest)
            {
                footer = "Clicking Yes will also Delete All the Contacts from the Change Tracking Operations in CRM (Enterprise Edition Only). If you are not sure about this, please Cancel and review the data in CRM.";
                icon = MessageBoxIcon.Asterisk;
            }

            DialogResult dr = ShowMessage("Start Synchronization Process", "Are you sure you want to start the Synchronization Process?", footer, MessageBoxButtons.YesNo, icon);

            if (dr == DialogResult.Yes)
            {
                if (!isTest)
                {
                    if (License.LicenseType.ToLower() == "enterprise")
                    {
                        tslCurrent.Text = "Deleting Contacts from Change Tracking";
                        Trace.AddLog(EventLevel.Information, DateTime.Now, "ToolClick", "General", "Start Executing Change Tracking Operations", "");
                        SyncCore.ExecuteChangeTrackingOperations();
                        Trace.AddLog(EventLevel.Information, DateTime.Now, "ToolClick", "General", "End Executing Change Tracking Operations", "");
                    }
                }

                tspbStatus.Minimum = 0; tspbStatus.Value = 0;
                int totalrowcount = dgvSyncList.Rows.Count;
                tslCurrent.ForeColor = System.Drawing.Color.Black;

                int totalRows = dgvSyncList.Rows.Count;
                if (isTest)
                    totalRows = dgvSyncList.SelectedRows.Count;

                if (totalRows > 0)
                {
                    if (!isTest)
                        SyncCore.UpdateSyncStart();

                    for (int i = 0; i < totalRows; i++)
                    {
                        DataGridViewRow row = dgvSyncList.Rows[i];

                        if (isTest)
                            row = dgvSyncList.SelectedRows[i];

                        Guid syncListId = new Guid(row.Cells[SyncListColumn.SyncListId.ToInt()].Value.ToString());
                        Guid listId = new Guid(row.Cells[SyncListColumn.ObjectId.ToInt()].Value.ToString());
                        string entityname = row.Cells[SyncListColumn.ObjectTypeName.ToInt()].Value.ToString();

                        string name = row.Cells[SyncListColumn.ListName.ToInt()].Value.ToString();
                        string listType = row.Cells[SyncListColumn.ListType.ToInt()].Value.ToString();
                        string exchangeGroupName = row.Cells[SyncListColumn.ExchangeGroupName.ToInt()].Value.ToString();
                        string exchangeGroupType = row.Cells[SyncListColumn.GroupType.ToInt()].Value == null ? "" : row.Cells[SyncListColumn.GroupType.ToInt()].Value.ToString();
                        DateTime lastsync = row.Cells[SyncListColumn.LastSync.ToInt()].Value == null ? DateTime.MinValue : Convert.ToDateTime(row.Cells[SyncListColumn.LastSync.ToInt()].Value);
                        int status = Convert.ToInt32(row.Cells[SyncListColumn.SyncStatus.ToInt()].Value);
                        if (status == 1)
                        {
                            tslCurrent.Text = "Current: " + name + " (" + (i + 1).ToString() + "/" + totalrowcount.ToString() + ")";
                            System.Threading.Thread.Sleep(500); Application.DoEvents();

                            SyncCore.ExecuteSync(syncListId, listId, entityname, name, exchangeGroupName, lastsync, listType);
                        }
                        else
                        {
                            Trace.AddLog(EventLevel.Information, DateTime.Now, "RunSync", "CRM", "Inactive List " + name + " (" + (i + 1).ToString() + "/" + totalrowcount.ToString() + ") not processed", string.Empty);
                            tslCurrent.Text = "Inactive: " + name + " (" + (i + 1).ToString() + "/" + totalrowcount.ToString() + ")";
                            System.Threading.Thread.Sleep(150); Application.DoEvents();
                        }
                        System.Threading.Thread.Sleep(500); Application.DoEvents();
                    }
                    Trace.AddLog(EventLevel.Information, DateTime.Now, "RunSync", "General", "Synchronization Completed", string.Empty);

                    if (!isTest)
                        SyncCore.UpdateSyncEnd();

                    tslCurrent.Text = "Completed Sync";
                    tslCurrent.ForeColor = System.Drawing.Color.Blue;

                    // Refresh Grid to Show Latest Sync Times
                    PopulateSyncList();
                }
                else
                {
                    MessageBox.Show("You must select at least one list to Synchronize", "Missing Lists", MessageBoxButtons.OK);
                }
            }

        }

        private void ValidateSync(bool quickValidate = true)
        {
            if (utResults.Nodes.Count > 0)
                utResults.Nodes.Clear();

            int totalRows = dgvSyncList.Rows.Count;
            if (totalRows > 0)
            {
                DataSet ds = new DataSet();
                DataTable dt = new DataTable("Validation");
                DataColumn col = new DataColumn("Result", typeof(string));
                dt.Columns.AddRange(new DataColumn[] { col });
                ds.Tables.Add(dt);

                for (int i = 0; i < totalRows; i++)
                {
                    DataGridViewRow row = dgvSyncList.Rows[i];

                    Guid syncListId = new Guid(row.Cells[SyncListColumn.SyncListId.ToInt()].Value.ToString());
                    Guid listId = new Guid(row.Cells[SyncListColumn.ObjectId.ToInt()].Value.ToString());
                    string entityname = row.Cells[SyncListColumn.ObjectTypeName.ToInt()].Value.ToString();

                    string name = row.Cells[SyncListColumn.ListName.ToInt()].Value.ToString();
                    string exchangeGroupName = row.Cells[SyncListColumn.ExchangeGroupName.ToInt()].Value.ToString();
                    DateTime lastsync = row.Cells[SyncListColumn.LastSync.ToInt()].Value == null ? DateTime.MinValue : Convert.ToDateTime(row.Cells[SyncListColumn.LastSync.ToInt()].Value);
                    int status = Convert.ToInt32(row.Cells[SyncListColumn.SyncStatus.ToInt()].Value);
                    if (status == 1)
                    {
                        System.Threading.Thread.Sleep(250); Application.DoEvents();
                        Validator result = SyncCore.ValidateSync(syncListId, listId, entityname, name, exchangeGroupName);
                        utResults.Visible = true;
                        if (result.TotalCRMContacts != result.TotalExchangeRecipients)
                        {
                            // Mismatch
                            string listItemText = String.Format ("{0} mismatch. CRM:{1}/Exchange{2}", result.CRMListName, result.TotalCRMContacts, result.TotalExchangeRecipients);
                            // lstResults.Items.Add(listItemText);
                            Infragistics.Win.UltraWinTree.UltraTreeNode node = utResults.Nodes.Add(result.CRMListName, listItemText);
                            // node.Cells[0].Appearance.ForeColor = Color.Red;
                        }
                        else
                        {
                            // Match
                            string listItemText = String.Format("{0} matches Exchange Group {1} with {2} Recipients", result.CRMListName, result.ExchangeGroupName, result.TotalExchangeRecipients);
                            // lstResults.Items.Add(listItemText);
                            Infragistics.Win.UltraWinTree.UltraTreeNode node = utResults.Nodes.Add(result.CRMListName, listItemText);
                            // node.Cells[0].Appearance.ForeColor = Color.Red;

                        }
                    }
                    System.Threading.Thread.Sleep(500); Application.DoEvents();
                }
            }
        }

        private void CompareList()
        {
            if (dgvSyncList.SelectedRows.Count > 0)
            {
                DataGridViewRow row = dgvSyncList.SelectedRows[0];

                Guid syncListId = new Guid(row.Cells[SyncListColumn.SyncListId.ToInt()].Value.ToString());
                Guid listId = new Guid(row.Cells[SyncListColumn.ObjectId.ToInt()].Value.ToString());

                string name = row.Cells[SyncListColumn.ListName.ToInt()].Value.ToString();
                string entityName = row.Cells[SyncListColumn.ObjectTypeName.ToInt()].Value.ToString();
                string exchangeGroupName = row.Cells[SyncListColumn.ExchangeGroupName.ToInt()].Value.ToString();

                ListCompare frm = new ListCompare(listId, entityName, name, exchangeGroupName);
                frm.ShowDialog(this);
            }
            else
            {
                MessageBox.Show("You must select a list to run comparison feature", "Select List", MessageBoxButtons.OK);
            }
        }

        private void CompareAll()
        {
            ListCompare frm = new ListCompare();
            frm.ShowDialog(this);
        }

        private void DetectDuplicates()
        {
            DuplicateDetect frm = new DuplicateDetect();
            frm.ShowDialog();
        }


        private void ClearAllExchangeAliases()
        {
            tspbStatus.Minimum = 0; tspbStatus.Value = 0;

            EntityCollection contacts = SyncCore.crm.RetrieveContactsWithAlias();
            if (contacts.Entities.Count > 0)
            {
                tspbStatus.Maximum = contacts.Entities.Count + 1;
                foreach (Entity contact in contacts.Entities)
                {
                    Guid contactid = contact.Id;
                    int revision = contact.Contains("xrm_revision") ? Convert.ToInt32(contact.Attributes["xrm_revision"]) : 0;

                    SyncCore.crm.UpdateContactAlias(contactid, "", revision);
                    System.Threading.Thread.Sleep(125); Application.DoEvents();
                    tspbStatus.Value++;
                }
                tslCurrent.Text = "A total of " + contacts.Entities.Count + " alias fields cleared";
            }
            else
                tslCurrent.Text = "No alias fields containing data were found in CRM";
        }

        private EntityCollection RetrieveCustomEntityInfo(string entityname, Guid id)
        {
            XmlDocument doc = new XmlDocument();
            doc.Load("Intersects.xml");

            XmlNode root = doc.DocumentElement;
            XmlNode node = root.SelectSingleNode("//Entity[@Name='"+ entityname + "']");
            string idFieldName = node.SelectSingleNode("//IdFieldName").InnerText;

            XmlNode child = node.SelectSingleNode("//IntersectEntity");
            string intersectType = child.SelectSingleNode("//Type").InnerText;
            string intersectName = child.SelectSingleNode("//Name").InnerText;
            string intersectParentIdField = child.SelectSingleNode("//ParentIdField").InnerText;
            string intersectContactIdField = child.SelectSingleNode("//ContactIdField").InnerText;

            EntityCollection contacts = new EntityCollection();
            if (intersectType.ToLower() == "manual")
            {
                contacts = SyncCore.crm.RetrieveManualIntersectMembers(intersectName, intersectContactIdField, intersectParentIdField, id);
            }
            else if (intersectType.ToLower() == "native")
            {
                contacts = SyncCore.crm.RetrieveNativeIntersectMembers(entityname, idFieldName, id, intersectName, intersectContactIdField, intersectParentIdField);
            }
            return contacts;
        }

        private void RetrieveXmlFile()
        {
            XmlDocument doc = new XmlDocument();
            doc.Load("Intersects.xml");

            XmlNode root = doc.DocumentElement;
            XmlNodeList list = root.SelectNodes("//Entity");

            foreach (XmlNode node in list)
            {
                string entityName = node.Attributes["Name"].Value;
                string schemaName = node.SelectSingleNode("//SchemaName").InnerText;
                string idFieldName = node.SelectSingleNode("//IdFieldName").InnerText;
                string primaryAttributeName = node.SelectSingleNode("//PrimaryAttributeName").InnerText;
                string DisplayName = node.SelectSingleNode("//DisplayName").InnerText;

                XmlNode child = node.SelectSingleNode("//IntersectEntity");
                string intersectType = child.SelectSingleNode("//Type").InnerText;
                string intersectName = child.SelectSingleNode("//Name").InnerText;
                string intersectParentIdField = child.SelectSingleNode("//ParentIdField").InnerText;
                string intersectContactIdField = child.SelectSingleNode("//ContactIdField").InnerText;
            }
        }

        private void AddMultiple()
        {
            ListManager frm = new ListManager();
            DialogResult dr = frm.ShowDialog(this);
            if (dr == DialogResult.OK)
            {
                // Refresh Sync List
                PopulateSyncList(true);
            }
        }


        private void AddEditGroup(bool isNew)
        {
            frmAddNew frm = new frmAddNew();
            if (isNew)
            {
                frm.FormType = frmAddNew.FormTypeCode.Create;
                frm.ShowDialog(this);
                PopulateSyncList();
            }
            else
            {
                if (dgvSyncList.SelectedRows.Count == 1)
                {
                    DataGridViewRow row = dgvSyncList.SelectedRows[0];
                    frm.FormType = frmAddNew.FormTypeCode.Update;
                    frm.SyncListId = new Guid(row.Cells[SyncListColumn.SyncListId.ToInt()].Value.ToString());
                    frm.EntityLogicalName = row.Cells[SyncListColumn.ObjectTypeName.ToInt()].Value.ToString();
                    frm.EntityName = row.Cells[SyncListColumn.ObjectType.ToInt()].Value.ToString();
                    frm.EntityDisplayName = row.Cells[SyncListColumn.ListName.ToInt()].Value.ToString();
                    frm.EntityId = new Guid(row.Cells[SyncListColumn.ObjectId.ToInt()].Value.ToString());
                    frm.ExchangeListName = row.Cells[SyncListColumn.ExchangeGroupName.ToInt()].Value.ToString();
                    frm.ShowDialog(this);
                    PopulateSyncList();

                }
                else if (dgvSyncList.SelectedRows.Count > 1)
                {
                    ShowMessage("Select Row", "Please select only one row when Editing a Group Connection", string.Empty, MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                }
                else // 0 rows selected 
                {
                    // Display Message. Must select row to Edit
                    ShowMessage("Select Row", "Please select a row if you want to Edit a Group Connection", string.Empty, MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
        }

        private void RemoveGroup()
        {
            int selectedRows = dgvSyncList.SelectedRows.Count;

            if (selectedRows > 0)
            {
                DialogResult rc = ShowMessage("Remove Selected Rows", "Are you sure you want to remove the selected rows from the Sync?", "Removing the rows does not remove the Lists and Groups from CRM and Exchange", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (rc == DialogResult.Yes)
                {
                    foreach (DataGridViewRow row in dgvSyncList.SelectedRows)
                    {
                        Guid rowId = new Guid(row.Cells[SyncListColumn.SyncListId.ToInt()].Value.ToString());
                        string entityname = row.Cells[SyncListColumn.ObjectTypeName.ToInt()].Value.ToString();

                        if (entityname != "")
                        {
                            SyncCore.crm.RemoveFromSyncList(rowId);
                        }
                    }
                    PopulateSyncList(true);
                }
            }
            else
            {
                ShowMessage("Select Row(s)", "Please select at least one row by clicking on the row selectors on the left", string.Empty, MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void ChangeStatus(bool enabled)
        {
            int selectedRows = dgvSyncList.SelectedRows.Count;
            if (selectedRows > 0)
            {
                foreach (DataGridViewRow row in dgvSyncList.SelectedRows)
                {
                    Guid rowId = new Guid(row.Cells[SyncListColumn.SyncListId.ToInt()].Value.ToString());
                    string entityname = row.Cells[SyncListColumn.ObjectTypeName.ToInt()].Value.ToString();

                    if (entityname != "")
                    {
                        SyncCore.crm.ChangeSyncListStatus(rowId, enabled);
                    }
                }
                PopulateSyncList(true);
            }
            else
            {
                ShowMessage("No Rows Selected", "Please select a row or rows by clicking on the row selectors on the left", "", MessageBoxButtons.OK, MessageBoxIcon.None);
            }
        }

        private void ChangeStatusAll(bool enabled)
        {
            for (int i = 0; i < dgvSyncList.Rows.Count; i++)
            {
                DataGridViewRow row = dgvSyncList.Rows[i];
                Guid rowId = new Guid(row.Cells[SyncListColumn.SyncListId.ToInt()].Value.ToString());
                string entityname = row.Cells[SyncListColumn.ObjectTypeName.ToInt()].Value.ToString();

                if (entityname != "")
                {
                    SyncCore.crm.ChangeSyncListStatus(rowId, enabled);
                }
            }
            PopulateSyncList();            
        }

        private void ChangeDataGridDisplay(bool showAll)
        {
            PopulateSyncList(!showAll);
        }

        private void ExportExchangeGroups()
        {
            DialogResult dr = MessageBox.Show("This process will Clear any information currently in the Log Cache. \nAre you sure you want to clear the log and run the Export Process?", "Export Exchange Groups", MessageBoxButtons.YesNo);
            if (dr == DialogResult.Yes)
            {
                List<DistributionGroup> groups = SyncCore.exch.GetAllDistributionListGroups();
                if (groups.Count > 0)
                {
                    ExcelExchangeGroups excel = new ExcelExchangeGroups();
                    Workbook book = excel.CreateBook();
                    foreach (DistributionGroup group in groups)
                    {
                        string groupName = group.Name;
                        List<ExchangeContact> contacts = SyncCore.exch.GetDistributionListMembers(groupName);
                        foreach (ExchangeContact contact in contacts)
                        {
                            string name = contact.Name;
                            string emailAddress = contact.PrimarySmtpAddress;
                            string alias = contact.Alias;
                            string identity = contact.Identity;
                            string ou = contact.OrganizationalUnit;
                            excel.AddRow(book, groupName, name, emailAddress, alias, identity, ou);
                        }
                    }

                    if (!System.IO.Directory.Exists(@"c:\logs"))
                    {
                        System.IO.Directory.CreateDirectory(@"C:\logs");
                    }

                    excel.SaveBook(book);
                }
            }
        }

        #region Logging Methods

        private void ViewLog()
        {
            if (!Trace.Null)
            {
                if (!Trace.Empty)
                {
                    TraceViewer frm = new TraceViewer();
                    frm.ShowDialog(this);
                }
                else
                {
                    ShowMessage("View Trace", "The Trace File is empty", string.Empty, MessageBoxButtons.OK, MessageBoxIcon.None);
                }
            }
            else
            {
                ShowMessage("View Trace", "The Trace File is empty", string.Empty, MessageBoxButtons.OK, MessageBoxIcon.None);
            }
        }

        private void ExportLogFile()
        {
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

                foreach(LogRow row in list)
                {
                    excel.AddRow(book, row.Level, row.LogDateTime, row.Category, row.MethodName, row.CommandName, row.Message, row.Details, row.Parameters);
                }
                excel.SaveBook(book);
            }


            /*
            List<DistributionGroup> groups = SyncCore.exch.GetAllDistributionListGroups();
            if (groups.Count > 0)
            {
                ExcelExchangeGroups excel = new ExcelExchangeGroups();
                Workbook book = excel.CreateBook();
                foreach (DistributionGroup group in groups)
                {
                    string groupName = group.Name;
                    List<ExchangeContact> contacts = SyncCore.exch.GetDistributionListMembers(groupName);
                    foreach (ExchangeContact contact in contacts)
                    {
                        string name = contact.Name;
                        string emailAddress = contact.PrimarySmtpAddress;
                        string alias = contact.Alias;
                        string identity = contact.Identity;
                        string ou = contact.OrganizationalUnit;
                        excel.AddRow(book, groupName, name, emailAddress, alias, identity, ou);
                    }
                }
                excel.SaveBook(book);
            }
             * */
        }

        #endregion

        private void tslConnection_DoubleClick(object sender, EventArgs e)
        {
            
        }

        
        private DialogResult ShowMessage(string captionText, string bodyText, string footerText, MessageBoxButtons buttons, MessageBoxIcon icon)
        {
            UltraMessageBoxInfo info = new UltraMessageBoxInfo();
            info.Caption = captionText;
            info.Text = bodyText;
            if (!string.IsNullOrEmpty(footerText))
                info.FooterFormatted = footerText;

            info.Buttons = buttons;
            info.Icon = icon;
            DialogResult dr = UltraMessageBoxManager.Show(info);
            return dr;
        }

        private void btnChangeProfile_Click(object sender, EventArgs e)
        {
            Profile frm = new Profile();
            DialogResult dr = frm.ShowDialog(this);
            if (dr == DialogResult.OK)
            {
                // Change Profile Configuration
            }
        }

        private void ExchangeSyncManager_Shown(object sender, EventArgs e)
        {
            try
            {
                if (ActiveForm != null)
                {
                    ActiveForm.Width = ActiveForm.Width + 10;
                    ActiveForm.Height = ActiveForm.Height + 10;
                }
            }
            catch (System.Exception ex)
            {

            }
        }

        private void ExchangeSyncManager_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (SyncCore != null)
            {
                if (SyncCore.exch != null)
                    SyncCore.exch.Disconnect();

                if (SyncCore.crm != null)
                    SyncCore.crm.Disconnect();
            }
            Application.Exit();
        }

        private void dgvSyncList_RowHeaderMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            System.Diagnostics.Debug.WriteLine("Selected Rows: " + dgvSyncList.SelectedRows.Count.ToString());
        }

        private void ExchangeSyncManager_Resize(object sender, EventArgs e)
        {
            if (ActiveForm != null)
            {
                int width = ActiveForm.Width;
                int height = ActiveForm.Height;

                dgvSyncList.Width = width - panel1.Width;
                dgvSyncList.Height = Form1_Fill_Panel.Height;
            }
        }

        private void pbPanelClose_Click(object sender, EventArgs e)
        {
            panel1.Dock = DockStyle.None;
            panel1.Width = 0;
            panel1.Visible = false;

            if (ActiveForm != null)
                dgvSyncList.Width = ActiveForm.Width;
        }
    }
}
