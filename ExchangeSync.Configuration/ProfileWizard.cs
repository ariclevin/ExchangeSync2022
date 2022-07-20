using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Windows.Forms;
using System.Xml;
using System.Xml.Linq;
using Microsoft.Xrm.Sdk;

namespace ExchangeSync
{
    public partial class ProfileWizard : Form
    {
        public string AppConfigFileName { get; set; }
        public string AppDir { get; set; }
        private int currentStep = 1;

        public string ProfileName { get; set; }
        Core syncCore;

        public ProfileWizard()
        {
            InitializeComponent();

            string filename = System.Reflection.Assembly.GetExecutingAssembly().Location;
            AppDir = Path.GetDirectoryName(filename);
            AppConfigFileName = AppDir + @"\CRMExchangeSync.exe";
        }

        public ProfileWizard(string profileName)
        {
            InitializeComponent();

            string filename = System.Reflection.Assembly.GetExecutingAssembly().Location;
            AppDir = Path.GetDirectoryName(filename);
            AppConfigFileName = AppDir + @"\CRMExchangeSync.exe";

            ProfileName = profileName;
        }

        private void frmRegisterNew_Load(object sender, EventArgs e)
        {
            syncCore = new Core();
            ShowPage(1);
        }

        private void ShowPage(int pageNo)
        {
            utRegistrationWizard.SelectedTab = utRegistrationWizard.Tabs[pageNo - 1];
            switch (pageNo)
            {
                case 1:
                    lblSectionTitle.Text = "Enter Profile Name";
                    btnPrevious.Enabled = false;
                    btnNext.Enabled = false;
                    btnRegister.Text = "Close";
                    break;
                case 2:
                    lblSectionTitle.Text = "CRM Connection Information";
                    btnPrevious.Enabled = true;
                    btnNext.Enabled = false;
                    btnRegister.Text = "Close";
                    break;
                case 3:
                    lblSectionTitle.Text = "Active Directory/Exchange Configuration";
                    btnPrevious.Enabled = true;
                    btnNext.Enabled = false;
                    btnRegister.Text = "Close";
                    break;
                case 4:
                    lblSectionTitle.Text = "Configuration Summary";
                    btnPrevious.Enabled = true;
                    btnNext.Enabled = false;
                    btnRegister.Text = "Finish";
                    DisplaySummary();
                    break;
            }
        }

        #region Click Events

        private void btnRegister_Click(object sender, EventArgs e)
        {
            if (btnRegister.Text == "Close")
            {
                DialogResult dr = MessageBox.Show("Are you sure you want to exit the Profile Creation Process?", "Exit", MessageBoxButtons.YesNo);
                if (dr == DialogResult.Yes)
                    Application.Exit();
            }
            else if (btnRegister.Text == "Finish")
            {
                // Remove Old Version Values from Application Settings
                CompleteConfiguration();
            }
            else if (btnRegister.Text == "Exit")
            {
                // syncCore.DisconnectFromCRM();
                // syncCore.DisconnectFromExchange();
                Application.Exit();
            }
        }


        private void CompleteConfiguration()
        {
            // Copy Template As New Config File

            string filename = System.Reflection.Assembly.GetExecutingAssembly().Location;
            string appDir = Path.GetDirectoryName(filename);

            string profileId = Guid.NewGuid().ToString();

            string sourceFile = appDir + @"\ProfileConfiguration.xml";
            string targetFile = appDir + @"\" + profileId + ".config";
            
            try
            {
                File.Copy(sourceFile, targetFile);
            }
            catch (System.Exception ex)
            {

            }

            // Open Configuration File
            Configuration config;
            ExeConfigurationFileMap configFileMap = new ExeConfigurationFileMap();
            configFileMap.ExeConfigFilename = targetFile;
            config = ConfigurationManager.OpenMappedExeConfiguration(configFileMap, ConfigurationUserLevel.None);
            KeyValueConfigurationCollection settings = config.AppSettings.Settings;

            if (settings.AllKeys.Contains<string>("ProfileName"))
                settings["ProfileName"].Value = txtProfileName.Text;
            else
                settings.Add("ProfileName", txtProfileName.Text);
            Trace.AddLog(EventLevel.Information, DateTime.Now, "UpdateConfigurationFile", "Add/Edit Setting", "Add/Editing Profile Name", string.Empty);
            Application.DoEvents(); Thread.Sleep(100);

            if (settings.AllKeys.Contains<string>("ProfileDesc"))
                settings["ProfileDesc"].Value = txtProfileDesc.Text;
            else
                settings.Add("ProfileDesc", txtProfileDesc.Text);
            Trace.AddLog(EventLevel.Information, DateTime.Now, "UpdateConfigurationFile", "Add/Edit Setting", "Add/Editing Profile Description", string.Empty);
            Application.DoEvents(); Thread.Sleep(100);


            if (settings.AllKeys.Contains<string>("CRMServiceUrl"))
                settings["CRMServiceUrl"].Value = txtCrmServiceUrl.Text;
            else
                settings.Add("CRMServiceUrl", txtCrmServiceUrl.Text);
            Trace.AddLog(EventLevel.Information, DateTime.Now, "UpdateConfigurationFile", "Add/Edit Setting", "Add/Editing Setting CRMServiceUrl", string.Empty);
            Application.DoEvents(); Thread.Sleep(100);

            if (settings.AllKeys.Contains<string>("CRMAuthenticationProviderType"))
                settings["CRMAuthenticationProviderType"].Value = (cmbAuthenticationType.SelectedIndex + 1).ToString();
            else
                settings.Add("CRMAuthenticationProviderType", (cmbAuthenticationType.SelectedIndex + 1).ToString());
            Trace.AddLog(EventLevel.Information, DateTime.Now, "UpdateConfigurationFile", "Add/Edit Setting", "Add/Editing Setting CRMAuthenticationProviderType", string.Empty);
            Application.DoEvents(); Thread.Sleep(100);

            if (settings.AllKeys.Contains<string>("CRMUsername"))
                settings["CRMUsername"].Value = txtCRMUsername.Text;
            else
                settings.Add("CRMUsername", txtCRMUsername.Text);
            Trace.AddLog(EventLevel.Information, DateTime.Now, "UpdateConfigurationFile", "Add/Edit Setting", "Add/Editing Setting CRMUsername", string.Empty);
            Application.DoEvents(); Thread.Sleep(100);

            if (settings.AllKeys.Contains<string>("CRMDomain"))
                settings["CRMDomain"].Value = txtCRMDomain.Text;
            else
                settings.Add("CRMDomain", txtCRMDomain.Text);
            Trace.AddLog(EventLevel.Information, DateTime.Now, "UpdateConfigurationFile", "Add/Edit Setting", "Add/Editing Setting CRMDomain", string.Empty);
            Application.DoEvents(); Thread.Sleep(100);

            if (settings.AllKeys.Contains<string>("ExchangeServerName"))
                settings["ExchangeServerName"].Value = txtExchangeServer.Text;
            else
                settings.Add("ExchangeServerName", txtExchangeServer.Text);
            Trace.AddLog(EventLevel.Information, DateTime.Now, "UpdateConfigurationFile", "Add/Edit Setting", "Add/Editing Setting ExchangeServerName", string.Empty);
            Application.DoEvents(); Thread.Sleep(100);

            string versionNumber = ConfigSetting.ExchangeServerVersion.ToInt().ToString();

            if (settings.AllKeys.Contains<string>("ExchangeServerVersion"))
                settings["ExchangeServerVersion"].Value = versionNumber;
            else
                settings.Add("ExchangeServerVersion", versionNumber);
            Trace.AddLog(EventLevel.Information, DateTime.Now, "UpdateConfigurationFile", "Add/Edit Setting", "Add/Editing Setting ExchangeServerVersion", string.Empty);
            Application.DoEvents(); Thread.Sleep(100);


            if (settings.AllKeys.Contains<string>("ADUsername"))
                settings["ADUsername"].Value = txtADUsername.Text;
            else
                settings.Add("ADUsername", txtADUsername.Text);
            Trace.AddLog(EventLevel.Information, DateTime.Now, "UpdateConfigurationFile", "Add/Edit Setting", "Add/Editing Setting ADUsername", string.Empty);
            Application.DoEvents(); Thread.Sleep(100);

            if (settings.AllKeys.Contains<string>("ADDomain"))
                settings["ADDomain"].Value = txtADDomain.Text;
            else
                settings.Add("ADDomain", txtADDomain.Text);
            Trace.AddLog(EventLevel.Information, DateTime.Now, "UpdateConfigurationFile", "Add/Edit Setting", "Add/Editing Setting ADDomain", string.Empty);
            Application.DoEvents(); Thread.Sleep(100);

            config.Save(ConfigurationSaveMode.Modified);
            Trace.AddLog(EventLevel.Information, DateTime.Now, "UpdateConfigurationFile", "Save Configuration", "Saving Configuration Changes", string.Empty);
            ConfigurationManager.RefreshSection("appSettings");

            ConfigurationSection section = config.GetSection(ProgramSetting.EncryptedSectionName);

            ClientSettingsSection clientSection = (ClientSettingsSection)section;

            SettingElement ADPassword = new SettingElement("ADPassword", SettingsSerializeAs.String);
            XmlElement ADElement = new XmlDocument().CreateElement("value");
            ADElement.InnerText = txtADPassword.Text;
            ADPassword.Value.ValueXml = ADElement;
            clientSection.Settings.Add(ADPassword);
            Trace.AddLog(EventLevel.Information, DateTime.Now, "UpdateConfigurationFile", "Add/Edit Setting", "Add/Editing Setting ADPassword", string.Empty);

            SettingElement CRMPassword = new SettingElement("CRMPassword", SettingsSerializeAs.String);
            XmlElement CRMElement = new XmlDocument().CreateElement("value");
            CRMElement.InnerText = txtCRMPassword.Text;
            CRMPassword.Value.ValueXml = CRMElement;
            clientSection.Settings.Add(CRMPassword);
            Trace.AddLog(EventLevel.Information, DateTime.Now, "UpdateConfigurationFile", "Add/Edit Setting", "Add/Editing Setting CRMPassword", string.Empty);

            section.SectionInformation.ForceSave = true;
            config.Save(ConfigurationSaveMode.Modified);
            Trace.AddLog(EventLevel.Information, DateTime.Now, "UpdateConfigurationFile", "Save Configuration", "Saving Encrypted Configuration Changes", string.Empty);
            ConfigurationManager.RefreshSection(ProgramSetting.EncryptedSectionName);

            section = config.GetSection(ProgramSetting.EncryptedSectionName);

            if (!section.SectionInformation.IsProtected)
            {
                //Protecting the specified section with the specified provider
                section.SectionInformation.ProtectSection("DPAPIProtection");
            }

            section.SectionInformation.ForceSave = true;
            config.Save(ConfigurationSaveMode.Modified);
            ConfigurationManager.RefreshSection(ProgramSetting.EncryptedSectionName);

            if (chkImportAppSettings.Checked)
                ImportApplicationSettings();

            if (chkImportFieldMappings.Checked)
                ImportFieldMappings();

            btnNext.Enabled = false;
            btnPrevious.Enabled = false;
            btnCancel.Enabled = false;
            btnRegister.Text = "Exit";
            lblCompleteConfirmation.Text = "Your Profile has been saved. Please click Exit button and run the Exchange Sync App.";
        }

        private void ImportApplicationSettings()
        {
            List<ApplicationSetting> appSettings = Excel.OpenApplicationSettings("ApplicationSettings.xlsx");
            Trace.AddLog(EventLevel.Information, DateTime.Now, "ImportApplicationSettings", "Import Settings", "Reading Application Settings", string.Empty);
            int totalSettings = appSettings.Count;

            foreach (ApplicationSetting item in appSettings)
            {
                string key = item.Key;
                Entity entity = syncCore.crm.RetrieveApplicationSetting(key);
                if ((entity != null) && (entity.Id != Guid.Empty))
                {
                    // This application setting already exists in CRM
                    // syncCore.crm.UpdateApplicationSetting(new Guid(item.Key), item.Value);
                    // Trace.AddLog(EventLevel.Information, DateTime.Now, "ImportApplicationSettings", "Update Settings", "Update Application Setting " + key, string.Empty);
                }
                else
                {
                    syncCore.crm.CreateApplicationSetting(11, item.Key, item.Value, item.Description);
                    Trace.AddLog(EventLevel.Information, DateTime.Now, "ImportApplicationSettings", "Create Settings", "Create Application Setting " + key, string.Empty);
                }
                Application.DoEvents(); Thread.Sleep(100);
            }
            
        }

        private void ImportFieldMappings()
        {

            List<FieldMapping> fieldMappings = Excel.OpenFieldMappings("FieldMappings.xlsx");
            int totalSettings = fieldMappings.Count;

            // Check if already connected???
            Uri connectUrl = new Uri(txtCrmServiceUrl.Text);
            int authenticationType = cmbAuthenticationType.SelectedIndex + 1;


            foreach (FieldMapping item in fieldMappings)
            {
                string crmFieldName = item.CRMFieldName;
                Entity entity = syncCore.crm.RetrieveFieldMap(crmFieldName);
                if ((entity != null) && (entity.Id != Guid.Empty))
                {
                    // This application setting already exists in CRM
                }
                else
                {
                    syncCore.crm.CreateFieldMapping(item.CRMFieldName, item.ExchangeFieldName, item.FieldType, item.DependencyType);
                    Trace.AddLog(EventLevel.Information, DateTime.Now, "ImportFieldMappings", "Create Field Mapping", "Create Field Mapping " + crmFieldName, string.Empty);
                }
                Application.DoEvents(); Thread.Sleep(100);
            }
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void btnPrevious_Click(object sender, EventArgs e)
        {
            ShowPage(currentStep - 1);
            currentStep--;
        }

        private void btnNext_Click(object sender, EventArgs e)
        {
            ShowPage(currentStep + 1);
            currentStep++;
        }


        private void btnTestConnection_Click(object sender, EventArgs e)
        {
            int authenticationType = cmbAuthenticationType.SelectedIndex + 1;
            if (ValidateCRMConnectionFields(authenticationType))
            {
                if (authenticationType != 3)
                {
                    aiCRMConnect.Visible = true; aiCRMConnect.Start();
                    CRMEventArgs args = syncCore.ConnectToCRM(ConfigSetting.AuthenticationProvider.ToInt(), ConfigSetting.CRMCredentials, ConfigSetting.CRMServiceUrl, ConfigSetting.OrganizationName, ConfigSetting.RegionName);

                    if (args.Result == true)
                    {
                        aiCRMConnect.Stop(); aiCRMConnect.Visible = false;

                        lblCRMConnectionMessage.Text = "Connection to CRM was successful";
                        lblCRMConnectionMessage.Visible = true;
                        btnNext.Enabled = true;

                    }
                    else
                    {
                        lblCRMConnectionMessage.Text = "Could not retrieve logged in user information from CRM. Connection unsuccessful.";
                        lblCRMConnectionMessage.ForeColor = System.Drawing.Color.Red;
                        lblCRMConnectionMessage.Visible = true;
                    }
                }
            }
            else
            {
                MessageBox.Show("Live Id (Former CRM Online) has been deprecated. Please select another Authentication method");
                btnNext.Enabled = false;
            }

            aiCRMConnect.Stop(); aiCRMConnect.Visible = false;
        }

        private bool ValidateCRMConnectionFields(int authenticationType)
        {
            bool rc = true;

            if (string.IsNullOrEmpty(txtCrmServiceUrl.Text))
            {
                epWizard.SetError(txtCrmServiceUrl, "You must specify the CRM Organization Service Url");
                rc = false;
            }

            switch (authenticationType)
            {
                case 1: // Active Directory
                    break;
                case 2: // Federated
                        if (string.IsNullOrWhiteSpace(txtCRMUsername.Text))
                        {
                            epWizard.SetError(txtCRMUsername, "You must specify a username to connect to CRM");
                            rc = false;
                        }

                        if (string.IsNullOrWhiteSpace(txtCRMPassword.Text))
                        {
                            epWizard.SetError(txtCRMPassword, "You must enter a password to connect to CRM");
                            rc = false;
                        }
                    break;
                case 3: // CRM Live (Deprecated)
                    rc = false;
                    break;
                case 4: // Office 365 CRM Online
                        if (string.IsNullOrWhiteSpace(txtCRMUsername.Text))
                        {
                            epWizard.SetError(txtCRMUsername, "You must specify a username to connect to CRM");
                            rc = false;
                        }

                        if (string.IsNullOrWhiteSpace(txtCRMPassword.Text))
                        {
                            epWizard.SetError(txtCRMPassword, "You must enter a password to connect to CRM");
                            rc = false;
                        }
                    break;
            }

            return rc;
        }

        private void btnTestExchangeConnection_Click(object sender, EventArgs e)
        {
            bool hasErrors = false;
            bool enabled = false;

            aiExchangeConnect.Start(); aiExchangeConnect.Visible = true;

            ExchangeServerType exchangeVersion = ExchangeServerType.Undefined;
            switch (cmbExchangeServerVersion.SelectedItem.ToString())
            {
                case "2010":
                    enabled = true;
                    exchangeVersion = ExchangeServerType.Exchange_2010;
                    break;
                case "2013":
                    enabled = true;
                    exchangeVersion = ExchangeServerType.Exchange_2013;
                    break;
                case "2016":
                    enabled = true;
                    exchangeVersion = ExchangeServerType.Exchange_2016;
                    break;
                case "Online":
                    enabled = false;
                    exchangeVersion = ExchangeServerType.Exchange_Online_365;
                    break;
            }

            if (string.IsNullOrEmpty(txtADUsername.Text))
            {
                epWizard.SetError(txtADUsername, "The username field is required");
                hasErrors = true;
            }

            if (string.IsNullOrEmpty(txtADPassword.Text))
            {
                epWizard.SetError(txtADPassword, "The password field is required");
                hasErrors = true;
            }

            if (enabled)
            {
                if (string.IsNullOrEmpty(txtADDomain.Text))
                {
                    epWizard.SetError(txtADDomain, "The domain name field is required");
                    hasErrors = true;
                }

                if (string.IsNullOrEmpty(txtExchangeServer.Text))
                {
                    epWizard.SetError(txtExchangeServer, "The Exchange Server name is required");
                    hasErrors = true;
                }
            }

            if (string.IsNullOrEmpty(cmbExchangeServerVersion.Text))
            {
                if (cmbExchangeServerVersion.Text != "Online")
                {
                    epWizard.SetError(cmbExchangeServerVersion, "The Exchange Server version number is required");
                    hasErrors = true;
                }
            }

            if (!hasErrors)
            {
                string userName = txtADUsername.Text;
                string password = txtADPassword.Text;
                string domain = txtADDomain.Text;
                string serverName = txtExchangeServer.Text;

                NetworkCredential credential = new NetworkCredential(userName, password, domain);
                ExchangeConnectEventArgs rc = syncCore.ConnectExchange(serverName, exchangeVersion, credential);

                if (rc.Result == true)
                {
                    lblTestExchangeConnectionConfirmation.Text = "Connected to Exchange";
                    lblTestExchangeConnectionConfirmation.ForeColor = Color.Black;
                    btnNext.Enabled = true;
                }
                else
                {
                    lblTestExchangeConnectionConfirmation.Text = "Connection to Exchange Failed.";
                    lblTestExchangeConnectionConfirmation.ForeColor = Color.Red;
                }
            }

            aiExchangeConnect.Stop(); aiExchangeConnect.Visible = false;
        }


        #endregion

        private bool ValidateUrl(string url, bool requiresHttp)
        {
            bool rc = true;
            if (requiresHttp)
            {
                if (!url.StartsWith("http"))
                {
                    url = "http://" + url;
                }

                if (!url.ToLower().EndsWith("/xrmservices/2011/organization.svc"))
                    rc = false;
            }
            else
            {
                if (url.StartsWith("http"))
                {
                    rc = false;
                }
            }

            if (rc)
            {
                rc = url.IsValidUrl();
            }

            return rc;
        }

        private bool ValidateDomainName(string domainName)
        {
            bool rc = true;

            if (!domainName.Contains('.'))
                rc = false;
            return rc;

        }

        private void cmbAuthenticationType_SelectedIndexChanged(object sender, EventArgs e)
        {
            switch (cmbAuthenticationType.SelectedIndex)
            {
                case 0: // On-Premise 
                    txtCRMDomain.ReadOnly = true; txtCRMDomain.Enabled = false;
                    txtCRMUsername.ReadOnly = true; txtCRMUsername.Enabled = false;
                    txtCRMPassword.ReadOnly = true; txtCRMPassword.Enabled = false;
                    break;
                case 1: // Federated (On-Premise, Partner Hosted or CRM Online)
                    txtCRMDomain.Text = ""; txtCRMDomain.ReadOnly = true; txtCRMDomain.Enabled = false;
                    txtCRMUsername.ReadOnly = false; txtCRMUsername.Enabled = true;
                    txtCRMPassword.ReadOnly = false; txtCRMPassword.Enabled = true;
                    break;
                case 2: // CRM Online Legacy
                    txtCRMDomain.Text = ""; txtCRMDomain.ReadOnly = true; txtCRMDomain.Enabled = false;
                    txtCRMUsername.Text = ""; txtCRMUsername.ReadOnly = true; txtCRMUsername.Enabled = false;
                    txtCRMPassword.Text = ""; txtCRMPassword.ReadOnly = true; txtCRMPassword.Enabled = false;
                    txtCrmServiceUrl.Enabled = false;
                    MessageBox.Show("Live Id (Former CRM Online) has been deprecated. Please select another Authentication method");
                    break;
                case 3: // Office 365
                    txtCRMDomain.Text = ""; txtCRMDomain.ReadOnly = true; txtCRMDomain.Enabled = false;
                    txtCRMUsername.ReadOnly = false; txtCRMUsername.Enabled = true;
                    txtCRMPassword.ReadOnly = false; txtCRMPassword.Enabled = true;
                    break;
            }
        }

        private void DisplaySummary()
        {
            StringBuilder sb = new StringBuilder();
            sb.AppendLine("Your Profile will be created/updated with the following parameters:");
            sb.AppendLine();

            sb.AppendLine();
            sb.AppendLine(String.Format("{0}: {1}", "Profile Name", txtProfileName.Text));
            
            sb.AppendLine();
            sb.AppendLine(String.Format("{0}: {1}", "CRM Authentication Method", cmbAuthenticationType.SelectedItem));
            sb.AppendLine(String.Format("{0}: {1}", "CRM Server Url", txtCrmServiceUrl.Text));

            sb.AppendLine();
            sb.AppendLine(String.Format("{0}: {1}", "Exchange Username", txtADUsername.Text));
            sb.AppendLine(String.Format("{0}: {1}", "Exchange Domain", txtADDomain.Text));
            sb.AppendLine(String.Format("{0}: {1}", "Exchange Server Version", cmbExchangeServerVersion.SelectedItem));

            txtSummary.Text = sb.ToString();
        }

        private void cmbExchangeServerVersion_SelectedIndexChanged(object sender, EventArgs e)
        {
            bool enabled = false;
            ExchangeServerType version = ExchangeServerType.Undefined;
            switch (cmbExchangeServerVersion.SelectedItem.ToString())
            {
                case "2010":
                    enabled = true;
                    version = ExchangeServerType.Exchange_2010;
                    break;
                case "2013":
                    enabled = true;
                    version = ExchangeServerType.Exchange_2013;
                    break;
                case "2016":
                    enabled = true;
                    version = ExchangeServerType.Exchange_2016;
                    break;
                case "Online":
                    txtExchangeServer.Text = string.Empty;
                    txtADDomain.Text = string.Empty;
                    version = ExchangeServerType.Exchange_Online_365;
                    enabled = false;
                    break;
            }
            txtExchangeServer.Enabled = enabled;
            txtADDomain.Enabled = enabled;
            ConfigSetting.ExchangeServerVersion = version;
        }

        private void label3_Click(object sender, EventArgs e)
        {

        }

        private void btnLoadDefault_Click(object sender, EventArgs e)
        {
            syncCore.LoadConfigurationSettings();

            txtCrmServiceUrl.Text = ConfigSetting.CRMServiceUrl;
            txtCRMUsername.Text = ConfigSetting.CRMCredentials.UserName;
            txtCRMDomain.Text = ConfigSetting.CRMCredentials.Domain;
            if (ConfigSetting.AuthenticationProvider.ToInt() != 0)
                cmbAuthenticationType.SelectedIndex = ConfigSetting.AuthenticationProvider.ToInt() - 1;

            // AD/Exchange Prepopulated Information
            txtADUsername.Text = ConfigSetting.ExchangeCredentials.UserName;
            txtADDomain.Text = ConfigSetting.ExchangeCredentials.Domain;
            txtExchangeServer.Text = ConfigSetting.ExchangeServerUrl;

            switch (ConfigSetting.ExchangeServerVersion)
            {
                case ExchangeServerType.Exchange_2010:
                    cmbExchangeServerVersion.SelectedIndex = 0;
                    break;
                case ExchangeServerType.Exchange_2013:
                    cmbExchangeServerVersion.SelectedIndex = 1;
                    break;
                case ExchangeServerType.Exchange_2016:
                    cmbExchangeServerVersion.SelectedIndex = 2;
                    break;
                case ExchangeServerType.Exchange_Online_365:
                    cmbExchangeServerVersion.SelectedIndex = 3;
                    break;

            }
        }

        private void txtProfileName_KeyUp(object sender, KeyEventArgs e)
        {
            if (txtProfileName.Text.Length >= 3)
            {
                btnNext.Enabled = true;
            }
        }

        private void label2_Click(object sender, EventArgs e)
        {

        }
    }


}
