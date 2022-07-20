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

using Infragistics.Documents.Excel;
using Microsoft.Xrm.Sdk;

namespace ExchangeSync
{
    public partial class ConfiguratorWizard : Form
    {
        public string AppConfigFileName { get; set; }
        public string AppDir { get; set; }
        const string SOLUTION_VERSION = "5.0.0.0";
        private int currentStep = 1;

        Guid licenseKeyId = Guid.Empty;
        System.Exception currentException;
        string locale = "en-US"; // default US locale
        bool saveUserInfo = false;

        public Core syncCore;
        ExchangeServerType versionType = ExchangeServerType.Undefined;

        public ConfiguratorWizard()
        {
            InitializeComponent();

            string filename = System.Reflection.Assembly.GetExecutingAssembly().Location;
            AppDir = Path.GetDirectoryName(filename);
            AppConfigFileName = AppDir + @"\ExchangeSync.exe";
        }

        private void frmRegisterNew_Load(object sender, EventArgs e)
        {
            syncCore = new Core();
            bool licenseExists = syncCore.LoadLicensingInformation();

            if (licenseExists)
            {
                string machineName = Network.FQDN.ToLower();
                string machineName2 = License.ComputerName.ToLower();

                if (machineName == machineName2)
                {
                    SetConfigurationKeys();
                    PresetFields();
                    // client = new XrmService.XrmServiceClient("basicHttp");
                    lblMachineInfo.Text = "Application will be installed and configured on " + machineName2;
                    lblMachineInfo.Visible = true;
                }
                else
                {
                    lblMachineInfo.Text = "Application can only be installed on " + machineName2;
                    lblMachineInfo.ForeColor = System.Drawing.Color.Red;
                    lblMachineInfo.Visible = true;
                    btnNext.Enabled = false;
                    btnValidateLicenseKey.Enabled = false;
                }

                ShowPage(1);
            }
            else
            {


                MessageBox.Show("A license file does not exist. Please install license file and run the configuration wizard again", "Missing License File", MessageBoxButtons.OK);
                Application.Exit();
            }
        }

        private void SetConfigurationKeys()
        {
            syncCore.LoadConfigurationSettings();

        }

        private void PresetFields()
        {
        // CRM Prepopulated Information
            txtCrmServiceUrl.Text = ConfigSetting.CRMServiceUrl;
            txtCRMUsername.Text = ConfigSetting.CRMCredentials.UserName;
            txtCRMDomain.Text = ConfigSetting.CRMCredentials.Domain;
            txtOrganizationName.Text = ConfigSetting.OrganizationName;
            cmbRegion.Text = ConfigSetting.RegionName;

            if (ConfigSetting.AuthenticationProvider.ToInt() != 0)
                cmbAuthenticationType.SelectedIndex = ConfigSetting.AuthenticationProvider.ToInt() - 1;
            chkIntegrated.Checked = ConfigSetting.IntegratedAuthentication;

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

            cmbDistributionGroupOU.Text = AppSetting.ExchangeDistributionGroupOU.Value;
            cmbDistributionGroupOU.Enabled = false;

            cmbContactsOU.Text = AppSetting.ExchangeMailContactOU.Value;
            cmbContactsOU.Enabled = false;

            txtExternalDomainName.Text = AppSetting.ExternalDomainName.Value;
        }

        private void RetrievePersonalInformation()
        {
            if (!string.IsNullOrEmpty(License.CompanyName))
            {
                txtCompany.Text = License.CompanyName;
                txtCompany.Enabled = false;
            }

            if (!string.IsNullOrEmpty(License.CompanyWebSite))
            {
                txtWebSiteUrl.Text = License.CompanyWebSite;
                txtWebSiteUrl.Enabled = false;
            }

            if (!string.IsNullOrEmpty(License.RegisteredToFirstName))
                txtFirstName.Text = License.RegisteredToFirstName;

            if (!string.IsNullOrEmpty(License.RegisteredToLastName))
                txtLastName.Text = License.RegisteredToLastName;

            if (!string.IsNullOrEmpty(License.RegisteredToPhoneNumber))
                txtPhoneNumber.Text = License.RegisteredToPhoneNumber;

            if (!string.IsNullOrEmpty(License.RegisteredToEmailAddress))
                txtEmailAddress.Text = License.RegisteredToEmailAddress;



        }

        private void CreatePersonalInformation()
        {

        }

        private void ShowPage(int pageNo)
        {
            utRegistrationWizard.SelectedTab = utRegistrationWizard.Tabs[pageNo - 1];
            switch (pageNo)
            {
                case 1:
                    lblSectionTitle.Text = "License Key";
                    btnPrevious.Enabled = false;
                    btnNext.Enabled = false;
                    btnRegister.Text = "Close";

                    switch (License.LicenseType.ToLower())
                    {
                        case "evaluation":
                            txtSerial.Enabled = false;
                            cmbEdition.SelectedIndex = 0;
                            btnValidateLicenseKey.Enabled = false;
                            btnNext.Enabled = true;
                            break;
                        case "community":
                            txtSerial.Enabled = false;
                            cmbEdition.SelectedIndex = 1;
                            btnValidateLicenseKey.Enabled = false;
                            btnNext.Enabled = true;
                            break;
                        case "enterprise":
                            Security crypt = new Security();
                            if (crypt.ValidateKey(License.LicenseKey))
                            {
                                txtSerial.Text = License.LicenseKey;
                                txtSerial.Enabled = false;
                                cmbEdition.SelectedIndex = 2;
                            }
                            break;
                        default:
                            txtSerial.Enabled = false;
                            cmbEdition.SelectedIndex = 0;
                            btnValidateLicenseKey.Enabled = false;
                            btnNext.Enabled = true;
                            break;
                    }
                    break;
                case 2:
                    lblSectionTitle.Text = "Customer Information";
                    btnPrevious.Enabled = true;
                    btnNext.Enabled = false;
                    btnRegister.Text = "Close";
                    RetrievePersonalInformation();
                    ValidateCustomerInformation();
                    break;
                case 3: // CRM Configuration
                    lblSectionTitle.Text = "CRM Connection Information";
                    btnPrevious.Enabled = true;
                    btnNext.Enabled = false;
                    btnRegister.Text = "Close";
                    aiCRMConnect.Visible = false;
                    SetClientSecretDisplayMode(false);

                    if (License.LicenseType.ToLower() == "community")
                    {
                        lblCRMPageTitle.Text = "Specify Connection to Dynamics 365 Online";
                        cmbAuthenticationType.SelectedIndex = 3;
                        cmbAuthenticationType.Enabled = false;

                        txtCRMDomain.Text = ""; txtCRMDomain.ReadOnly = true; txtCRMDomain.Enabled = false;
                        lblCRMDomain.Visible = false; txtCRMDomain.Visible = false;
                        txtCRMUsername.ReadOnly = false; txtCRMUsername.Enabled = true;
                        txtCRMPassword.ReadOnly = false; txtCRMPassword.Enabled = true;
                        txtOrganizationName.ReadOnly = false; txtOrganizationName.Enabled = true;
                        cmbRegion.Enabled = true;
                    }
                    break;
                case 4: // Exchange Configuration
                    lblSectionTitle.Text = "Active Directory/Exchange Configuration";
                    btnPrevious.Enabled = true;
                    btnNext.Enabled = false;
                    btnRegister.Text = "Close";
                    aiExchangeConnect.Visible = false;

                    if (License.LicenseType.ToLower() == "community")
                    {
                        lblCRMPageTitle.Text = lblCRMPageTitle.Text.Replace("Server", "Online");
                        cmbExchangeServerVersion.SelectedIndex = 3;
                        cmbExchangeServerVersion.Enabled = false;
                        txtExchangeServer.Enabled = false;

                        txtADDomain.Text = ""; txtADDomain.ReadOnly = true; txtADDomain.Enabled = false;
                        txtADDomain.Visible = false;
                        txtADUsername.ReadOnly = false; txtADUsername.Enabled = true;
                        txtADPassword.ReadOnly = false; txtADPassword.Enabled = true;
                        cmbDistributionGroupOU.Enabled = false;
                        cmbContactsOU.Enabled = false;
                    }

                    break;
                case 5:
                    lblSectionTitle.Text = "Configuration Summary";
                    btnPrevious.Enabled = true;
                    btnNext.Enabled = false;
                    btnRegister.Text = "Finish";
                    DisplaySummary();
                    break;
                case 6:
                    lblSectionTitle.Text = "Completing Configuration";
                    btnPrevious.Enabled = false;
                    btnNext.Enabled = false;
                    btnRegister.Text = "Exit";
                    break;
            }
        }

        private void FixConfigurationFile()
        {
            lblCurrentStep.Text = "Removing Legacy Application Settings";

            Configuration config = ConfigurationManager.OpenExeConfiguration(AppConfigFileName);
            KeyValueConfigurationCollection settings = config.AppSettings.Settings;

            int totalSettings = settings.Count;
            pbCurrent.Value = 0;

            foreach (KeyValueConfigurationElement element in settings)
            {
                switch (element.Key.ToLower())
                {
                    case "dlou":
                        settings.Remove(element.Key);
                        break;
                    //case "contactsou":
                    //    settings.Remove(element.Key);
                    //    break;
                    case "username":
                        settings.Remove(element.Key);
                        break;
                    case "liveidusername":
                        settings.Remove(element.Key);
                        break;
                    case "version":
                        settings.Remove(element.Key);
                        break;
                    case "customfieldname":
                        settings.Remove(element.Key);
                        break;
                    case "requiredattributes":
                        settings.Remove(element.Key);
                        break;
                    case "crmaliasfieldname":
                        settings.Remove(element.Key);
                        break;
                    case "crmrevisionfieldname":
                        settings.Remove(element.Key);
                        break;
                }
                Trace.AddLog(EventLevel.Information, DateTime.Now, "FixConfigurationFile", "Remove Setting", "Removing Setting " + element.Key.ToLower(), string.Empty);
                pbCurrent.Value += (100 / totalSettings);
                Application.DoEvents(); Thread.Sleep(100);
            }

            config.Save(ConfigurationSaveMode.Modified);
            Trace.AddLog(EventLevel.Information, DateTime.Now, "FixConfigurationFile", "Save Configuration", "Saving Configuration File", string.Empty);

            ConfigurationManager.RefreshSection("appSettings");
            pbCurrent.Value = 100;

            Application.DoEvents(); Thread.Sleep(500);

            lblCurrentStep.Text = "Removing Encrypted Legacy Application Settings";
            pbCurrent.Value = 0;

            ConfigurationSection section = config.GetSection(ProgramSetting.EncryptedSectionName);
            if (section.SectionInformation.IsProtected)
            {
                //Protecting the specified section with the specified provider
                section.SectionInformation.UnprotectSection();
            }

            section.SectionInformation.ForceSave = true;
            config.Save(ConfigurationSaveMode.Modified);
            Trace.AddLog(EventLevel.Information, DateTime.Now, "FixConfigurationFile", "Save Configuration", "Saving Configuration File (Unprotected)", string.Empty);

            // Remove All Protected Section Information
            ClientSettingsSection clientSection = (ClientSettingsSection)section;
            totalSettings = clientSection.Settings.Count;
            clientSection.Settings.Clear();

            for (int i = 0; i < totalSettings; i++)
            {
                pbCurrent.Value += (100 / totalSettings);
                Application.DoEvents(); Thread.Sleep(100);
            }

            config.Save(ConfigurationSaveMode.Modified);
            Trace.AddLog(EventLevel.Information, DateTime.Now, "FixConfigurationFile", "Save Configuration", "Saving Configuration File (Protected)", string.Empty);
            ConfigurationManager.RefreshSection(ProgramSetting.EncryptedSectionName);

            pbCurrent.Value = 100;

        }

        private void UpdateConfigurationFile()
        {
            lblCurrentStep.Text = "Updating Application Settings";

            int totalSettings = 13;
            pbCurrent.Value = 0;

            Configuration config = ConfigurationManager.OpenExeConfiguration(AppConfigFileName);
            KeyValueConfigurationCollection settings = config.AppSettings.Settings;

            if (settings.AllKeys.Contains<string>("CRMServiceUrl"))
                settings["CRMServiceUrl"].Value = txtCrmServiceUrl.Text; 
            else
                settings.Add("CRMServiceUrl", txtCrmServiceUrl.Text);
            Trace.AddLog(EventLevel.Information, DateTime.Now, "UpdateConfigurationFile", "Add/Edit Setting", "Add/Editing Setting CRMServiceUrl" , string.Empty);
            pbCurrent.Value += (100 / totalSettings); Application.DoEvents(); Thread.Sleep(100);

            if (settings.AllKeys.Contains<string>("ExchangeServerName"))
                settings["ExchangeServerName"].Value = txtExchangeServer.Text;
            else
                settings.Add("ExchangeServerName", txtExchangeServer.Text);
            Trace.AddLog(EventLevel.Information, DateTime.Now, "UpdateConfigurationFile", "Add/Edit Setting", "Add/Editing Setting ExchangeServerName", string.Empty);
            pbCurrent.Value += (100 / totalSettings); Application.DoEvents(); Thread.Sleep(100);

            if (settings.AllKeys.Contains<string>("ExchangeServerVersion"))
                settings["ExchangeServerVersion"].Value = cmbExchangeServerVersion.SelectedItem.ToString();
            else
                settings.Add("ExchangeServerVersion", cmbExchangeServerVersion.SelectedItem.ToString());
            Trace.AddLog(EventLevel.Information, DateTime.Now, "UpdateConfigurationFile", "Add/Edit Setting", "Add/Editing Setting ExchangeServerVersion", string.Empty); pbCurrent.Value += (100 / totalSettings); Application.DoEvents(); Thread.Sleep(100);

            if (settings.AllKeys.Contains<string>("DistributionGroupOU"))
            {
                if (cmbDistributionGroupOU.SelectedItem != null)
                    settings["DistributionGroupOU"].Value = cmbDistributionGroupOU.SelectedItem.ToString();
                else
                    settings["DistributionGroupOU"].Value = string.Empty;
            }
            else
            {
                if (cmbDistributionGroupOU.SelectedItem != null)
                    settings.Add("DistributionGroupOU", cmbDistributionGroupOU.SelectedItem.ToString());
                else
                    settings.Add("DistributionGroupOU", string.Empty);
            }
            Trace.AddLog(EventLevel.Information, DateTime.Now, "UpdateConfigurationFile", "Add/Edit Setting", "Add/Editing Setting DistributionGroupOU", string.Empty); pbCurrent.Value += (100 / totalSettings); Application.DoEvents(); Thread.Sleep(100);

            if (settings.AllKeys.Contains<string>("ContactsOU"))
            {
                if (cmbContactsOU.SelectedItem != null)
                    settings["ContactsOU"].Value = cmbContactsOU.SelectedItem.ToString();
                else
                    settings["ContactsOU"].Value = string.Empty;
            }
            else
            {
                if (cmbContactsOU.SelectedItem != null)
                    settings.Add("ContactsOU", cmbContactsOU.SelectedItem.ToString());
                else
                    settings.Add("ContactsOU", string.Empty);
            }
            Trace.AddLog(EventLevel.Information, DateTime.Now, "UpdateConfigurationFile", "Add/Edit Setting", "Add/Editing Setting ContactsOU", string.Empty); pbCurrent.Value += (100 / totalSettings); Application.DoEvents(); Thread.Sleep(100);

            if (settings.AllKeys.Contains<string>("CRMUsername"))
                settings["CRMUsername"].Value = txtCRMUsername.Text;
            else
                settings.Add("CRMUsername", txtCRMUsername.Text);
            Trace.AddLog(EventLevel.Information, DateTime.Now, "UpdateConfigurationFile", "Add/Edit Setting", "Add/Editing Setting CRMUsername", string.Empty); pbCurrent.Value += (100 / totalSettings); Application.DoEvents(); Thread.Sleep(100);

            if (settings.AllKeys.Contains<string>("CRMDomain"))
                settings["CRMDomain"].Value = txtCRMDomain.Text;
            else
                settings.Add("CRMDomain", txtCRMDomain.Text);
            Trace.AddLog(EventLevel.Information, DateTime.Now, "UpdateConfigurationFile", "Add/Edit Setting", "Add/Editing Setting CRMDomain", string.Empty); pbCurrent.Value += (100 / totalSettings); Application.DoEvents(); Thread.Sleep(100);

            if (settings.AllKeys.Contains<string>("CRMAuthenticationProviderType"))
                settings["CRMAuthenticationProviderType"].Value = (cmbAuthenticationType.SelectedIndex + 1).ToString();
            else
                settings.Add("CRMAuthenticationProviderType", (cmbAuthenticationType.SelectedIndex + 1).ToString());
            Trace.AddLog(EventLevel.Information, DateTime.Now, "UpdateConfigurationFile", "Add/Edit Setting", "Add/Editing Setting CRMAuthenticationProviderType", string.Empty); 
            pbCurrent.Value += (100 / totalSettings); Application.DoEvents(); Thread.Sleep(100);

            if (settings.AllKeys.Contains<string>("OrganizationName"))
                settings["OrganizationName"].Value = txtOrganizationName.Text;
            else
                settings.Add("OrganizationName", txtOrganizationName.Text);
            Trace.AddLog(EventLevel.Information, DateTime.Now, "UpdateConfigurationFile", "Add/Edit Setting", "Add/Editing Setting OrganizationName", string.Empty);
            pbCurrent.Value += (100 / totalSettings); Application.DoEvents(); Thread.Sleep(100);

            if (settings.AllKeys.Contains<string>("RegionName"))
                settings["RegionName"].Value = cmbRegion.Text;
            else
                settings.Add("RegionName", cmbRegion.Text);
            Trace.AddLog(EventLevel.Information, DateTime.Now, "UpdateConfigurationFile", "Add/Edit Setting", "Add/Editing Setting RegionName", string.Empty);
            pbCurrent.Value += (100 / totalSettings); Application.DoEvents(); Thread.Sleep(100);

            if (settings.AllKeys.Contains<string>("CRMIntegratedAuthentication"))
                settings["CRMIntegratedAuthentication"].Value = chkIntegrated.Checked ? "1" : "0";
            else
                settings.Add("CRMIntegratedAuthentication", chkIntegrated.Checked ? "1" : "0");
            Trace.AddLog(EventLevel.Information, DateTime.Now, "UpdateConfigurationFile", "Add/Edit Setting", "Add/Editing Setting CRMIntegratedAuthentication", string.Empty);
            pbCurrent.Value += (100 / totalSettings); Application.DoEvents(); Thread.Sleep(100);

            if (settings.AllKeys.Contains<string>("ADUsername"))
                settings["ADUsername"].Value = txtADUsername.Text;
            else
                settings.Add("ADUsername", txtADUsername.Text);
            Trace.AddLog(EventLevel.Information, DateTime.Now, "UpdateConfigurationFile", "Add/Edit Setting", "Add/Editing Setting ADUsername", string.Empty); 
            pbCurrent.Value += (100 / totalSettings); Application.DoEvents(); Thread.Sleep(100);

            if (settings.AllKeys.Contains<string>("ADDomain"))
                settings["ADDomain"].Value = txtADDomain.Text;
            else
                settings.Add("ADDomain", txtADDomain.Text);
            Trace.AddLog(EventLevel.Information, DateTime.Now, "UpdateConfigurationFile", "Add/Edit Setting", "Add/Editing Setting ADDomain", string.Empty);
            pbCurrent.Value += (100 / totalSettings); Application.DoEvents(); Thread.Sleep(100);

            // Add Connection String Value to Config File
            if (chkUseConnectionString.Checked)
            {
                if (settings.AllKeys.Contains<string>("UseCRMConnectionString"))
                    settings["UseCRMConnectionString"].Value = "1";
                else
                    settings.Add("UseCRMConnectionString", "1");
                Trace.AddLog(EventLevel.Information, DateTime.Now, "UpdateConfigurationFile", "Add/Edit Setting", "Add/Editing Setting ADDomain", string.Empty);
                pbCurrent.Value += (100 / totalSettings); Application.DoEvents(); Thread.Sleep(100);

                ConnectionStringsSection csSection = config.ConnectionStrings;
                ConnectionStringSettingsCollection csCollection = csSection.ConnectionStrings;
                ConnectionStringSettings cs = csCollection["Dynamics"];

                if (cs == null)
                    csCollection.Add(new ConnectionStringSettings("Dynamics", txtConnectionString.Text));
                else
                    csCollection["Dynamics"].ConnectionString = txtConnectionString.Text;
            }
            else
            {
                if (settings.AllKeys.Contains<string>("UseCRMConnectionString"))
                    settings["UseCRMConnectionString"].Value = "0";
                else
                    settings.Add("UseCRMConnectionString", "0");
                Trace.AddLog(EventLevel.Information, DateTime.Now, "UpdateConfigurationFile", "Add/Edit Setting", "Add/Editing Setting ADDomain", string.Empty);
                pbCurrent.Value += (100 / totalSettings); Application.DoEvents(); Thread.Sleep(100);
            }

            config.Save(ConfigurationSaveMode.Modified);
            Trace.AddLog(EventLevel.Information, DateTime.Now, "UpdateConfigurationFile", "Save Configuration", "Saving Configuration Changes", string.Empty);
            ConfigurationManager.RefreshSection("appSettings");
            pbCurrent.Value = 100;

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


        }

        private void SetCRMSecurityRole()
        {
            lblCurrentStep.Text = "Adding Security";

            int totalSettings = 1;
            pbCurrent.Value = 0;

            // Uri connectUrl = new Uri(txtCrmServiceUrl.Text);
            // int authenticationType = cmbAuthenticationType.SelectedIndex + 1;

            try
            {
                syncCore.ConnectToCRM(ConfigSetting.AuthenticationProvider.ToInt(), ConfigSetting.CRMCredentials, ConfigSetting.CRMServiceUrl, ConfigSetting.OrganizationName, ConfigSetting.RegionName);

                Guid userId = syncCore.crm.WhoAmI();
                Guid roleId = syncCore.crm.RetrieveRole("Exchange Sync");
                syncCore.crm.AssociateUserToRole(userId, roleId);
                Trace.AddLog(EventLevel.Information, DateTime.Now, "SetCRMSecurityRole", "Set Security Role", "Setting Security Role to CRM User", string.Empty);
            }
            catch (System.Exception ex)
            {
                Trace.AddLog(EventLevel.Warning, DateTime.Now, "SetCRMSecurityRole", "Set Security Role", "Setting Security Role to CRM User", "Security Role Already exists or an error has occured");
            }

            pbCurrent.Value = 100;

        }

        private void ImportApplicationSettings(string fileName)
        {
            
            pbCurrent.Value = 0;
            lblCurrentStep.Text = "Importing CRM Application Settings. Existing settings will not be updated";

            List<ApplicationSetting> appSettings = Excel.OpenApplicationSettings(fileName);
            Trace.AddLog(EventLevel.Information, DateTime.Now, "ImportApplicationSettings", "Import Settings", "Reading Application Settings", string.Empty); 
            int totalSettings = appSettings.Count;
            
            foreach (ApplicationSetting item in appSettings)
            {
                string key = item.Key;
                Entity entity = new Entity();

                try
                {
                    entity = syncCore.crm.RetrieveApplicationSetting(key);
                }
                catch (System.Exception ex)
                {
                    string message = ex.Message;
                    string command = "", function = "", parameterString = "";
                    if (ex.InnerException != null)
                    {
                        command = ex.InnerException.Data.Contains("CommandName") ? ex.InnerException.Data["CommandName"].ToString() : string.Empty;
                        function = ex.InnerException.Data.Contains("MethodName") ? ex.InnerException.Data["MethodName"].ToString() : string.Empty;

                        List<KeyValuePair<string, string>> parameters = ex.InnerException.Data.Contains("Parameters") ? (List<KeyValuePair<string, string>>)ex.InnerException.Data["Parameters"] : new List<KeyValuePair<string, string>>();
                        parameterString = Helper.KeyValuePairListToString(parameters);
                    }

                    Trace.AddLog(EventLevel.Error, DateTime.Now, "CRM", command, function, message, string.Empty, parameterString);
                }

                if (entity != null)
                {
                    // This application setting already exists in CRM
                    if (entity.Id != Guid.Empty)
                    {
                        switch (key)
                        {
                            case "ExchangeDistributionGroupOU":
                                if (cmbDistributionGroupOU.SelectedItem != null)
                                    UpdateApplicationSetting(entity.Id, cmbDistributionGroupOU.SelectedItem.ToString());
                                else
                                    UpdateApplicationSetting(entity.Id, string.Empty);
                                break;
                            case "ExchangeMailContactOU":
                                if (cmbContactsOU.SelectedItem != null)
                                    UpdateApplicationSetting(entity.Id, cmbContactsOU.SelectedItem.ToString());
                                else
                                    UpdateApplicationSetting(entity.Id, string.Empty);
                                break;
                            case "ExternalDomainName":
                                UpdateApplicationSetting(entity.Id, txtExternalDomainName.Text);
                                break;
                            default:
                                break;
                        }

                    }
                    Trace.AddLog(EventLevel.Information, DateTime.Now, "ImportApplicationSettings", "Update Settings", "Update Application Setting " + key, string.Empty); 

                }
                else
                {
                    switch (key)
                    {
                        case "ExchangeDistributionGroupOU":
                            if (cmbDistributionGroupOU.SelectedItem != null)
                                CreateApplicationSetting(11, item.Key, cmbDistributionGroupOU.SelectedItem.ToString(), item.Description);
                            else
                                CreateApplicationSetting(11, item.Key, string.Empty, item.Description);
                            break;
                        case "ExchangeMailContactOU":
                            if (cmbContactsOU.SelectedItem != null)
                                CreateApplicationSetting(11, item.Key, cmbContactsOU.SelectedItem.ToString(), item.Description);
                            else
                                CreateApplicationSetting(11, item.Key, string.Empty, item.Description);
                            break;
                        case "ExternalDomainName":
                            CreateApplicationSetting(11, item.Key, txtExternalDomainName.Text, item.Description);
                            break;
                        default:
                            CreateApplicationSetting(11, item.Key, item.Value, item.Description);
                            break;
                    }
                    Trace.AddLog(EventLevel.Information, DateTime.Now, "ImportApplicationSettings", "Create Settings", "Create Application Setting " + key, string.Empty);                     
                }
                pbCurrent.Value += (100 / totalSettings); Application.DoEvents(); Thread.Sleep(100);
            }
            pbCurrent.Value = 100;
        }

        private void CreateApplicationSetting(int categoryCode, string key, string value, string description)
        {
            try
            {
                syncCore.crm.CreateApplicationSetting(categoryCode, key, value, description);
            }
            catch (System.Exception ex)
            {
                string message = ex.Message;
                string command = ex.InnerException.Data["CommandName"].ToString();
                string function = ex.InnerException.Data["MethodName"].ToString();

                List<KeyValuePair<string, string>> parameters = (List<KeyValuePair<string, string>>)ex.InnerException.Data["Parameters"];
                string parameterString = Helper.KeyValuePairListToString(parameters);
                Trace.AddLog(EventLevel.Error, DateTime.Now, "CreateApplicationSetting", function, command, message, string.Empty, parameterString);
            }
        }

        private void UpdateApplicationSetting(Guid entityId, string value)
        {
            try
            {
                syncCore.crm.UpdateApplicationSetting(entityId, value);
            }
            catch (System.Exception ex)
            {
                string message = ex.Message;
                string command = ex.InnerException.Data["CommandName"].ToString();
                string function = ex.InnerException.Data["MethodName"].ToString();

                List<KeyValuePair<string, string>> parameters = (List<KeyValuePair<string, string>>)ex.InnerException.Data["Parameters"];
                string parameterString = Helper.KeyValuePairListToString(parameters);
                Trace.AddLog(EventLevel.Error, DateTime.Now, "CreateApplicationSetting", function, command, message, string.Empty, parameterString);
            }
        }


        private void ImportFieldMappings(string fileName)
        {

            pbCurrent.Value = 0;
            lblCurrentStep.Text = "Importing CRM Field Mappings. Existing mappings will not be updated";

            List<FieldMapping> fieldMappings = Excel.OpenFieldMappings(fileName);
            int totalSettings = fieldMappings.Count;

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
                    try
                    {
                        syncCore.crm.CreateFieldMapping(item.CRMFieldName, item.ExchangeFieldName, item.FieldType, item.DependencyType);
                        Trace.AddLog(EventLevel.Information, DateTime.Now, "ImportFieldMappings", "Create Field Mapping", "Create Field Mapping " + crmFieldName, string.Empty);
                    }
                    catch (System.Exception ex)
                    {
                        string message = ex.Message;
                        string command = ex.InnerException.Data["CommandName"].ToString();
                        string function = ex.InnerException.Data["MethodName"].ToString();

                        List<KeyValuePair<string, string>> parameters = (List<KeyValuePair<string, string>>)ex.InnerException.Data["Parameters"];
                        string parameterString = Helper.KeyValuePairListToString(parameters);
                        Trace.AddLog(EventLevel.Error, DateTime.Now, "ImportFieldMappings", function, command, message, string.Empty, parameterString);
                    }
                }
                pbCurrent.Value += (100 / totalSettings); Application.DoEvents(); Thread.Sleep(100);
            }
            pbCurrent.Value = 100;
        }
        

        #region Click Events

        private void btnRegister_Click(object sender, EventArgs e)
        {
            if (btnRegister.Text == "Close")
            {
                DialogResult dr = MessageBox.Show("Are you sure you want to exit the Registration Process?", "Exit", MessageBoxButtons.YesNo);
                if (dr == DialogResult.Yes)
                    Application.Exit();
            }
            else if (btnRegister.Text == "Finish")
            {
                // Remove Old Version Values from Application Settings
                ShowPage(6); currentStep++;
                CompleteConfiguration();
                txtCompleteConfigurationSummary.Visible = true;
            }
            else if (btnRegister.Text == "Exit")
            {
                ExportLogFile();
                Application.Exit();
            }
        }

        private void CompleteConfiguration()
        {
            if (Trace.Null)
                Trace.InitializeLog();
            else
            {
                if (!Trace.Empty)
                    Trace.ClearLog();
            }

            FixConfigurationFile();
            UpdateConfigurationFile();
            SetCRMSecurityRole();

            bool isConnected = false;
            try
            {
                if (syncCore.crm.conn != null)
                {
                    if (!syncCore.crm.conn.IsReady)
                    {
                        isConnected = Reconnect();
                    }
                    else
                        isConnected = true;
                }
                else
                {
                    isConnected = Reconnect();
                }
            }
            catch (System.Exception ex)
            {
                isConnected = Reconnect();
            }

            if (isConnected)
            {
                string currentPath = Path.GetDirectoryName(Application.ExecutablePath);
                currentPath += "\\Resources";

                FolderBrowserDialog dlg = new FolderBrowserDialog();
                dlg.RootFolder = Environment.SpecialFolder.MyComputer;
                dlg.SelectedPath = currentPath + "\\Resources";

                string fileName = string.Empty;
                if (chkImportAppSettings.Checked)
                {
                    fileName = currentPath + "\\ApplicationSettings.xlsx";
                    if (File.Exists(fileName))
                        ImportApplicationSettings(fileName);
                    else
                    {
                        DialogResult dr = dlg.ShowDialog(this);
                        if (dr == DialogResult.OK)
                        {
                            fileName = dlg.SelectedPath + "\\ApplicationSettings.xlsx";
                            ImportApplicationSettings(fileName);
                        }
                    }
                }

                if (chkImportFieldMappings.Checked)
                {
                    fileName = currentPath + "\\FieldMappings.xlsx";
                    if (File.Exists(fileName))
                        ImportFieldMappings(fileName);
                    else
                    {
                        DialogResult dr = dlg.ShowDialog(this);
                        if (dr == DialogResult.OK)
                        {
                            fileName = dlg.SelectedPath + "\\FieldMappings.xlsx";
                            ImportFieldMappings(fileName);
                        }
                    }

                }

            }
            else
            {
                MessageBox.Show("Connectivity to CRM has been lost. Import Application Settings and Field Mappings Manually", "CRM Connectivity List", MessageBoxButtons.OK);
            }
        }

        private bool Reconnect()
        {
            bool success = false;
            int authenticationType = cmbAuthenticationType.SelectedIndex + 1;
            if (authenticationType != 3)
            {
                NetworkCredential credentials = new NetworkCredential(txtCRMUsername.Text, txtCRMPassword.Text, txtCRMDomain.Text);
                CRMEventArgs args = syncCore.ConnectToCRM(authenticationType, credentials, txtCrmServiceUrl.Text, txtOrganizationName.Text, cmbRegion.Text);
                success = args.Result;
            }
            return success;        
        }

        private void ExportLogFile()
        {
            string filename = @"C:\logs\exchangesyncconfig.csv";
            System.IO.File.WriteAllText(filename, Trace.RetrieveLogAsText(true));
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


        private void btnDetails_Click(object sender, EventArgs e)
        {
            string machineName = Network.MachineName;
            if (machineName.Contains('.'))
                machineName = machineName.Substring(0, machineName.IndexOf('.'));
            // machineName = Program.GetMachineName(machineName);

            if (currentException != null)
                MessageBox.Show(currentException.Message, machineName);
            else
                MessageBox.Show("No Exception Available", machineName);
        }

        private void btnTestConnection_Click(object sender, EventArgs e)
        {
            int authenticationType = cmbAuthenticationType.SelectedIndex + 1;
            CRMEventArgs args;

            if (ValidateCRMConnectionFields(authenticationType))
            {
                if (authenticationType != 3)
                {
                    NetworkCredential credentials = new NetworkCredential(txtCRMUsername.Text, txtCRMPassword.Text, txtCRMDomain.Text);
                    aiCRMConnect.Visible = true; aiCRMConnect.Start();

                    if (chkUseConnectionString.Checked)
                        args = syncCore.ConnectToCRM(txtConnectionString.Text, txtCRMPassword.Text);
                    else
                        args = syncCore.ConnectToCRM(authenticationType, credentials, txtCrmServiceUrl.Text, txtOrganizationName.Text, cmbRegion.Text);
                }
                else
                {
                    string serverUrl = @"https://" + txtCrmServiceUrl.Text; // + @"/" + txtOrganizationName.Text;
                    args = syncCore.ConnectToCRM(serverUrl, txtClientId.Text, txtClientSecret.Text);
                }

                if (args.Result == true)
                {
                    aiCRMConnect.Stop(); aiCRMConnect.Visible = false;

                    lblCRMConnectionMessage.Text = "Connection to CRM was successful";
                    lblCRMConnectionMessage.Visible = true;

                    EntityCollection solutions = syncCore.crm.RetrieveSolutions("crmexchangesync");
                    if (solutions.Entities.Count > 0)
                    {
                        // Solution Exists
                        string friendlyName = solutions.Entities[0].Attributes["friendlyname"].ToString();
                        string versionNumber = solutions.Entities[0].Attributes["version"].ToString();

                        lblCRMSolutionMessage.Text = String.Format("{0} solution version {1} found in your CRM organization", friendlyName, versionNumber);
                        lblCRMSolutionMessage.Visible = true;
                        btnNext.Enabled = true;

                        // if (versionNumber != SOLUTION_VERSION)
                        if (versionNumber.ParseVersion() < SOLUTION_VERSION.ParseVersion())
                        {
                            lblCRMSolutionMessage.Text = String.Format("The version that you have installed is not the minimal required version. Please install an updated managed solution in order to continue.");
                            lblCRMSolutionMessage.ForeColor = System.Drawing.Color.Red;

                            btnInstallSolution.Text = "Upgrade Solution";
                            btnInstallSolution.Visible = true;
                            btnNext.Enabled = false;
                            chkSkipInstall.Visible = true;
                            chkSkipInstall.Text = "Manual solution upgrade";

                        }
                        else
                        {
                            btnInstallSolution.Text = "Upgrade Solution";
                            btnInstallSolution.Visible = true;
                            chkSkipInstall.Visible = true;
                            chkSkipInstall.Text = "Manual solution upgrade";
                        }
                    }
                    else
                    {
                        lblCRMSolutionMessage.Text = String.Format("No valid solution found in CRM. Please install the managed solution in order to continue.");
                        lblCRMSolutionMessage.ForeColor = System.Drawing.Color.Red;
                        lblCRMSolutionMessage.Visible = true;
                        btnNext.Enabled = false;
                        chkSkipInstall.Visible = true;

                        btnInstallSolution.Visible = true;
                    }
#if DEBUG
                    btnInstallSolution.Visible = true;
#endif
                }
                else
                {
                    aiCRMConnect.Stop(); aiCRMConnect.Visible = false;

                    lblCRMConnectionMessage.Text = "Could not retrieve logged in user information from CRM. Connection unsuccessful.";

                    Infragistics.Win.UltraWinToolTip.UltraToolTipInfo info = this.ttmWizard.GetUltraToolTip(this.lblCRMConnectionMessage);
                    info.ToolTipTitle = "Connection error";
                    info.ToolTipText = args.ConnectionReason + "\n" + args.LastCRMError;

                    lblCRMConnectionMessage.ForeColor = System.Drawing.Color.Red;
                    lblCRMConnectionMessage.Visible = true;
                }
            }
        }

        private void chkSkipInstall_CheckedChanged(object sender, EventArgs e)
        {
            btnNext.Enabled = true;
        }

        private bool ValidateCRMConnectionFields(int authenticationType)
        {
            bool rc = true;

            if (string.IsNullOrEmpty(txtCrmServiceUrl.Text))
            {
                epWizard.SetError(txtCrmServiceUrl, "You must specify the CRM Organization Server Address");
                rc = false;
            }
            else
            {
                if (License.LicenseType.ToLower() == "community")
                {
                    if (!txtCrmServiceUrl.Text.Contains("dynamics.com"))
                    {
                        epWizard.SetError(txtCrmServiceUrl, "Organization must be a CRM Online Organization for use with Community Edition");
                        rc = false;
                    }
                }
            }

            switch (authenticationType)
            {
                case 1: // Active Directory
                    if (!chkIntegrated.Checked)
                    {
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

                        if (string.IsNullOrWhiteSpace(txtCRMDomain.Text))
                        {
                            epWizard.SetError(txtCRMDomain, "You must specify an Active Directory domain to connect to CRM");
                            rc = false;
                        }

                        if (string.IsNullOrWhiteSpace(txtOrganizationName.Text))
                        {
                            epWizard.SetError(txtOrganizationName, "You must specify an Organization Name to connect to CRM");
                            rc = false;
                        }

                    }
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

                        if (string.IsNullOrWhiteSpace(txtOrganizationName.Text))
                        {
                            epWizard.SetError(txtOrganizationName, "You must specify an Organization Name to connect to CRM");
                            rc = false;
                        }

                    break;
                case 3: // oAuth (Not In Use)
                    if (string.IsNullOrWhiteSpace(txtClientId.Text))
                    {
                        epWizard.SetError(txtClientId, "You must specify a Client Id to connect to CRM using OAuth");
                        rc = false;
                    }

                    if (string.IsNullOrWhiteSpace(txtClientSecret.Text))
                    {
                        epWizard.SetError(txtClientSecret, "You must enter a Client Secret to connect to CRM using OAuth");
                        rc = false;
                    }
                    break;
                case 4: // CRM Online/CDS
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

                        if (string.IsNullOrWhiteSpace(txtOrganizationName.Text))
                        {
                            epWizard.SetError(txtOrganizationName, "You must specify an Organization Name to connect to CRM");
                            rc = false;
                        }

                        if (string.IsNullOrWhiteSpace(cmbRegion.Text))
                        {
                            epWizard.SetError(cmbRegion, "You must specify a Region to connect to CRM Online");
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

            switch (cmbExchangeServerVersion.SelectedItem.ToString())
            {
                case "2010":
                    enabled = true;
                    versionType = ExchangeServerType.Exchange_2010;
                    break;
                case "2013":
                    enabled = true;
                    versionType = ExchangeServerType.Exchange_2013;
                    break;
                case "2016":
                    enabled = true;
                    versionType = ExchangeServerType.Exchange_2016;
                    break;
                case "2019":
                    enabled = true;
                    versionType = ExchangeServerType.Exchange_2019;
                    break;
                case "Online":
                case "365":
                    enabled = false;
                    versionType = ExchangeServerType.Exchange_Online_365;
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
                // string userName = txtADUsername.Text;
                // string password = txtADPassword.Text;
                // string domain = txtADDomain.Text;
                // string serverName = txtExchangeServer.Text;

                NetworkCredential credentials = new NetworkCredential(txtADUsername.Text, txtADPassword.Text, txtADDomain.Text);
                aiExchangeConnect.Visible = true; aiExchangeConnect.Start();
                ExchangeConnectEventArgs args = syncCore.ConnectExchange(txtExchangeServer.Text, versionType, credentials);
                if (args.Result == true)
                {
                    if (versionType != ExchangeServerType.Exchange_Online_365)
                    {
                        string validName = syncCore.exch.GetServerName();

                        cmbDistributionGroupOU.Enabled = true; cmbContactsOU.Enabled = true;
                        List<OrganizationalUnit> groupList = syncCore.exch.GetAllOrganizationalUnits();
                        cmbDistributionGroupOU.Items.Clear(); cmbContactsOU.Items.Clear();

                        int currentIndex = 0;
                        foreach (OrganizationalUnit ou in groupList)
                        {
                            cmbDistributionGroupOU.Items.Add(ou.CanonicalName);
                            cmbContactsOU.Items.Add(ou.CanonicalName);

                            if (ou.CanonicalName == AppSetting.ExchangeDistributionGroupOU.Value)
                                cmbDistributionGroupOU.SelectedIndex = currentIndex;

                            if (ou.CanonicalName == AppSetting.ExchangeMailContactOU.Value)
                                cmbContactsOU.SelectedIndex = currentIndex;

                            currentIndex++;
                        }

                        if (!lblDistributionGroupOU.Text.StartsWith("Ex"))
                        {
                            cmbDistributionGroupOU.SelectedItem = lblDistributionGroupOU.Text;
                        }

                        if (!lblContactOU.Text.StartsWith("Ex"))
                        {
                            cmbContactsOU.SelectedItem = lblContactOU.Text;
                        }
                        
                        if (groupList.Count > 0)
                        {
                            lblTestExchangeConnectionConfirmation.Text = String.Format("Connected to {0} Exchange On-Premise {1} Server", validName, versionType.ToPlainString());
                            lblTestExchangeConnectionConfirmation.ForeColor = Color.Black;
                        }
                    }
                    else // versionType is ExchangeOnline
                    {
                        // Exchange Online -- NOT NECESSARY
                        // List<string> groupList = syncCore.exch.GetAllDistributionLists();
                        // if (groupList.Count > 0)
                        // {
                            lblTestExchangeConnectionConfirmation.Text = "Connected to Exchange Online";
                            lblTestExchangeConnectionConfirmation.ForeColor = Color.Black;
                            cmbDistributionGroupOU.Enabled = false;
                            cmbContactsOU.Enabled = false;
                        // }
                    }

                    aiExchangeConnect.Stop(); aiExchangeConnect.Visible = false;
                }
                else
                {
                    lblTestExchangeConnectionConfirmation.Text = "Unable to Connect to Exchange";
                    lblTestExchangeConnectionConfirmation.ForeColor = Color.Red;
                    cmbDistributionGroupOU.Enabled = false;
                    cmbContactsOU.Enabled = false;

                }
            }
        }

        private void btnInstallSolution_Click(object sender, EventArgs e)
        {
            OpenFileDialog dlg = new OpenFileDialog();
            dlg.InitialDirectory = System.IO.Path.GetDirectoryName(Application.ExecutablePath);
            dlg.Filter = "Zip Files (*.zip)|*.zip";
            dlg.Multiselect = false;

            DialogResult result = dlg.ShowDialog();

            if (result == DialogResult.OK)
            {
                string fileName = dlg.FileName;
                lblCRMSolutionMessage.Text = "Installing Solution, please wait..."; Thread.Sleep(100); Application.DoEvents();
                try
                {
                    syncCore.crm.ImportSolution(fileName);
                    lblCRMSolutionMessage.Text = String.Format("The CRM Solution {0} has been installed in your CRM Organization", fileName);
                    btnNext.Enabled = true;
                }
                catch (System.Exception ex)
                {
                    string message = ex.Message;
                    string command = ex.InnerException.Data.Contains("CommandName") ? ex.InnerException.Data["CommandName"].ToString() : string.Empty;
                    string function = ex.InnerException.Data.Contains("MethodName") ? ex.InnerException.Data["MethodName"].ToString() : string.Empty;

                    List<KeyValuePair<string, string>> parameters = ex.InnerException.Data.Contains("Parameters") ? (List<KeyValuePair<string, string>>)ex.InnerException.Data["Parameters"] : new List<KeyValuePair<string, string>>();
                    string parameterString = Helper.KeyValuePairListToString(parameters);

                    lblCRMSolutionMessage.Text = String.Format("Errors occured imported the CRM Solution {0} into your CRM Organization", fileName);
                    lblCRMSolutionMessage.ForeColor = Color.Red;

                    MessageBox.Show(String.Format("An error occured importing your solution. You will have to uninstall the existing solution and install the new solution manually. Click on the Backup button to export your existing lists. The error encountered was: {0}", ex.Message));
                    btnNext.Enabled = false;
                }
            }
        }

        private void btnRevalidateExchangeSettings_Click(object sender, EventArgs e)
        {
            bool showNext = true;

            int version = 9999;
            if (cmbExchangeServerVersion.Text != "Online")
                version = Convert.ToInt32(cmbExchangeServerVersion.Text);

            if (version != 9999)
            {
                if (cmbDistributionGroupOU.SelectedItem == null)
                    showNext = false;

                if (cmbContactsOU.SelectedItem == null)
                    showNext = false;
            }

            if (txtExternalDomainName.Text == "")
                showNext = false;

            btnNext.Enabled = showNext;
        }
        
        private void btnValidateLicenseKey_Click(object sender, EventArgs e)
        {
            lblLicenseConfirmation.Text = "";
            lblLicenseConfirmation.ForeColor = System.Drawing.Color.Black;

            if (!string.IsNullOrEmpty(txtSerial.Text))
            {
                if (txtSerial.Text == License.LicenseKey)
                {
                    lblLicenseConfirmation.Text = "Product Key valid. Click Next";
                    lblLicenseConfirmation.Visible = true;
                    epWizard.SetError(txtSerial, "");

                    // Validate Machine Name (Make sure not installed on multiple machines)
                    btnNext.Enabled = true;
                }
                else
                {
                    lblLicenseConfirmation.Text = "Invalid License Key. Please retry";
                    lblLicenseConfirmation.ForeColor = System.Drawing.Color.Red;
                    epWizard.SetError(txtSerial, "Invalid License Key. Please retry");
                    lblLicenseConfirmation.Visible = true;
                }
            }
            else
            {
                lblLicenseConfirmation.Text = "Product Key cannot be empty";
                lblLicenseConfirmation.ForeColor = System.Drawing.Color.Red;
                epWizard.SetError(txtSerial, "Product Key cannot be empty");
                lblLicenseConfirmation.Visible = true;
            }
        }

        #endregion

        private void ValidateCustomerInformation()
        {
            bool formValid = true;

            if (string.IsNullOrEmpty(txtCompany.Text))
            {
                if (formValid)
                    epWizard.SetError(txtCompany, "Company Name is a required field");
                formValid = false;
            }
            else
                epWizard.SetError(txtCompany, "");

            if (string.IsNullOrEmpty(txtWebSiteUrl.Text))
            {
                if (formValid)
                    epWizard.SetError(txtWebSiteUrl, "Web site is a required field");
                formValid = false;
            }
            else
            {
                if (ValidateUrl(txtWebSiteUrl.Text, false))
                    epWizard.SetError(txtWebSiteUrl, "");
                else
                {
                    if (formValid)
                        epWizard.SetError(txtWebSiteUrl, "The web site address you entered is not valid");
                    formValid = false;
                }
            }

            if (string.IsNullOrEmpty(txtPhoneNumber.Text))
            {
                if (formValid)
                    epWizard.SetError(txtPhoneNumber, "Phone Number is a required field");
                formValid = false;
            }
            else
                epWizard.SetError(txtPhoneNumber, "");

            if (string.IsNullOrEmpty(txtFirstName.Text))
            {
                if (formValid)
                    epWizard.SetError(txtFirstName, "First Name is a required field");
                formValid = false;
            }
            else
                epWizard.SetError(txtFirstName, "");

            if (string.IsNullOrEmpty(txtLastName.Text))
            {
                if (formValid)
                    epWizard.SetError(txtLastName, "Last Name is a required field");
                formValid = false;
            }
            else
                epWizard.SetError(txtLastName, "");

            if (string.IsNullOrEmpty(txtEmailAddress.Text))
            {
                if (formValid)
                    epWizard.SetError(txtEmailAddress, "Email Address is a required field");
                formValid = false;
            }
            else
            {
                if (txtEmailAddress.Text.IsValidEmail())
                    epWizard.SetError(txtEmailAddress, "");
                else
                {
                    if (formValid)
                        epWizard.SetError(txtEmailAddress, "The email address you entered is not valid");
                    formValid = false;
                }
            }

            if (formValid)
            {
                if (saveUserInfo)
                {
                    CreatePersonalInformation();
                    saveUserInfo = false;
                }
                btnNext.Enabled = true;
                btnNext.Focus();
            }
        }

        private void txtCompany_Leave(object sender, EventArgs e)
        {
            ValidateCustomerInformation();
        }

        private void txtCompany_TextChanged(object sender, EventArgs e)
        {
            saveUserInfo = true;
        }

        private void txtWebSiteUrl_Leave(object sender, EventArgs e)
        {
            ValidateCustomerInformation();
        }

        private void txtWebSiteUrl_TextChanged(object sender, EventArgs e)
        {
            saveUserInfo = true;
        }

        private void txtPhoneNumber_Leave(object sender, EventArgs e)
        {
            ValidateCustomerInformation();
        }

        private void txtPhoneNumber_TextChanged(object sender, EventArgs e)
        {
            saveUserInfo = true;
        }

        private void txtFirstName_Leave(object sender, EventArgs e)
        {
            ValidateCustomerInformation();
        }

        private void txtFirstName_TextChanged(object sender, EventArgs e)
        {
            saveUserInfo = true;
        }

        private void txtLastName_Leave(object sender, EventArgs e)
        {
            ValidateCustomerInformation();
        }

        private void txtLastName_TextChanged(object sender, EventArgs e)
        {
            saveUserInfo = true;
        }

        private void txtEmailAddress_Leave(object sender, EventArgs e)
        {
            ValidateCustomerInformation();
        }

        private void txtEmailAddress_TextChanged(object sender, EventArgs e)
        {
            saveUserInfo = true;
        }

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
                Regex RgxUrl = new Regex("(([a-zA-Z][0-9a-zA-Z+\\-\\.]*:)?/{0,2}[0-9a-zA-Z;/?:@&=+$\\.\\-_!~*'()%]+)?(#[0-9a-zA-Z;/?:@&=+$\\.\\-_!~*'()%]+)?");

                if (!RgxUrl.IsMatch(url))
                    rc = false;
            }

            return rc;
        }

        private bool ValidateEmailAddress(string emailAddress)
        {
            bool rc = true;
            Regex RgxEmail = new Regex(@"^(?("")(""[^""]+?""@)|(([0-9a-z]((\.(?!\.))|[-!#\$%&'\*\+/=\?\^`\{\}\|~\w])*)(?<=[0-9a-z])@))" +
                @"(?(\[)(\[(\d{1,3}\.){3}\d{1,3}\])|(([0-9a-z][-\w]*[0-9a-z]*\.)+[a-z0-9]{2,24}))$");

            if (!RgxEmail.IsMatch(emailAddress))
                rc = false;

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
                case 0: // AD (On-Premise)
                    // chkIntegrated.Enabled = true; chkIntegrated.Checked = true;
                    SetClientSecretDisplayMode(false);
                    lblSampleConnection.Text = "Ex. ==> crmserver:5555. If no port is entered, port 80 will be used.";
                    txtCRMDomain.ReadOnly = false; txtCRMDomain.Enabled = true;
                    txtCRMUsername.ReadOnly = false; txtCRMUsername.Enabled = true;
                    txtCRMPassword.ReadOnly = false; txtCRMPassword.Enabled = true;
                    break;
                case 1: // IFD (On-Premise, Partner Hosted)
                    // chkIntegrated.Checked = false; chkIntegrated.Enabled = false;
                    SetClientSecretDisplayMode(false);
                    lblSampleConnection.Text = "Ex. ==> crmserver.domain.com/orgname";
                    txtCRMDomain.Text = ""; txtCRMDomain.ReadOnly = true; txtCRMDomain.Enabled = false;
                    txtCRMUsername.ReadOnly = false; txtCRMUsername.Enabled = true;
                    txtCRMPassword.ReadOnly = false; txtCRMPassword.Enabled = true;
                    break;
                case 2: // OAuth
                    // chkIntegrated.Checked = false; chkIntegrated.Enabled = false;
                    SetClientSecretDisplayMode(true);
                    //txtCRMDomain.Text = ""; txtCRMDomain.ReadOnly = true; txtCRMDomain.Enabled = false;
                    //txtCRMUsername.Text = ""; txtCRMUsername.ReadOnly = true; txtCRMUsername.Enabled = false;
                    //txtCRMPassword.Text = ""; txtCRMPassword.ReadOnly = true; txtCRMPassword.Enabled = false;
                    //txtCrmServiceUrl.Enabled = false;
                    // MessageBox.Show("OAuth is currently under construction. Please use Alternate Method");
                    break;
                case 3: // Office 365
                    // chkIntegrated.Checked = false; chkIntegrated.Enabled = false;
                    SetClientSecretDisplayMode(false);
                    lblSampleConnection.Text = "Ex. ==> crmserver.crm.dynamics.com";
                    txtCRMDomain.Text = ""; txtCRMDomain.ReadOnly = true; txtCRMDomain.Enabled = false;
                    txtCRMUsername.ReadOnly = false; txtCRMUsername.Enabled = true;
                    txtCRMPassword.ReadOnly = false; txtCRMPassword.Enabled = true;
                    txtOrganizationName.ReadOnly = false; txtOrganizationName.Enabled = true;
                    cmbRegion.Enabled = true;
                    break;
            }
        }

        private void txtCrmServiceUrl_TextChanged(object sender, EventArgs e)
        {
            string serverUrl = txtCrmServiceUrl.Text;
            switch (cmbAuthenticationType.SelectedIndex)
            {
                case 0: // AD (On-Premise)
                    // Example: contoso:5555 or advworks:80
                    if (serverUrl.StartsWith("http") || serverUrl.Contains(@"/"))
                    {
                        epWizard.SetError(txtCrmServiceUrl, "Server Url should be in the format servername:port (contoso:5555), and not should not have http(s) prefix or Organization Service Url suffix");
                         
                    }
                    break;
                case 1: // IFD (On-Premise, Partner Hosted)
                    // chkIntegrated.Checked = false; chkIntegrated.Enabled = false;
                    // lblSampleConnection.Text = "Ex. ==> crmserver.domain.com/orgname";
                    if (serverUrl.StartsWith("http") || serverUrl.ToLower().Contains("xrmservices"))
                    {
                        epWizard.SetError(txtCrmServiceUrl, "Server Url should be in the format internalservername.domain.com, and not should not have http(s) prefix or Organization Service Url suffix");
                    }
                    break;
                case 2: // OAuth
                    // chkIntegrated.Checked = false; chkIntegrated.Enabled = false;
                    break;
                case 3: // Office 365
                        // chkIntegrated.Checked = false; chkIntegrated.Enabled = false;
                    if (serverUrl.StartsWith("http") || serverUrl.ToLower().Contains("xrmservices"))
                    {
                        epWizard.SetError(txtCrmServiceUrl, "Server Url should be in the format organization.crm#.dynamics.com, and not should not have http(s) prefix or Organization Service Url suffix");
                    }

                    break;
            }
        }

        private void chkIntegrated_CheckedChanged(object sender, EventArgs e)
        {
            if (chkIntegrated.Checked)
            {
                txtCRMUsername.Text = "";  txtCRMUsername.ReadOnly = true; txtCRMUsername.Enabled = false;
                txtCRMPassword.Text = ""; txtCRMPassword.ReadOnly = true; txtCRMPassword.Enabled = false;
                txtCRMDomain.Text = ""; txtCRMDomain.ReadOnly = true; txtCRMDomain.Enabled = false;
            }
            else
            {
                txtCRMUsername.ReadOnly = false; txtCRMUsername.Enabled = true;
                txtCRMPassword.ReadOnly = false; txtCRMPassword.Enabled = true;
                txtCRMDomain.ReadOnly = false; txtCRMDomain.Enabled = true;
            }
        }

        private void DisplaySummary()
        {
            StringBuilder sb = new StringBuilder();
            sb.AppendLine("Your CRM Exchange Sync will be created with the following parameters:");
            sb.AppendLine();

            sb.AppendLine(String.Format("Product Type: {0}", cmbEdition.Text));

            sb.AppendLine();
            sb.AppendLine(String.Format("{0}: {1}", "Company Name", txtCompany.Text));
            sb.AppendLine(String.Format("{0}: {1}", "Web Site", txtWebSiteUrl.Text));
            
            sb.AppendLine();
            sb.AppendLine(String.Format("{0}: {1}", "CRM Authentication Method", cmbAuthenticationType.SelectedItem));
            sb.AppendLine(String.Format("{0}: {1}", "CRM Server Url", txtCrmServiceUrl.Text));
            if (cmbAuthenticationType.SelectedIndex == 0)
            {
                if (chkIntegrated.Checked)
                {
                    sb.AppendLine(String.Format("{0}: {1}", "Integrated Authentication", "True"));
                }
                else
                {
                    sb.AppendLine(String.Format("{0}: {1}", "CRM Username", txtCRMUsername.Text));
                    sb.AppendLine(String.Format("{0}: {1}", "CRM Domain", txtCRMDomain.Text));
                }
            }
            else
            {
                sb.AppendLine(String.Format("{0}: {1}", "CRM Username", txtCRMUsername.Text));
            }

            sb.AppendLine();
            sb.AppendLine(String.Format("{0}: {1}", "Exchange Username", txtADUsername.Text));
            sb.AppendLine(String.Format("{0}: {1}", "Exchange Domain", txtADDomain.Text));
            sb.AppendLine(String.Format("{0}: {1}", "Exchange Server Version", cmbExchangeServerVersion.SelectedItem));

            if (cmbExchangeServerVersion.SelectedItem.ToString() != "Online")
            {
                sb.AppendLine(String.Format("{0}: {1}", "Exchange Server Name", txtExchangeServer.Text));
                sb.AppendLine(String.Format("{0}: {1}", "Distribution Group OU", cmbDistributionGroupOU.SelectedItem));
                sb.AppendLine(String.Format("{0}: {1}", "Mail Contacts OU", cmbContactsOU.SelectedItem));
            }
            else
            {
                sb.AppendLine("Server and Group information ommitted for Exchange Online");
            }
            txtSummary.Text = sb.ToString();
        }

        private string StringToHex(string input)
        {
            string rc = "";
            char[] values = input.ToCharArray();
            foreach (char letter in values)
            {
                int value = Convert.ToInt32(letter);
                string hexValue = String.Format("{0:X}", value);
                rc += hexValue;
            }
            return rc;
        }

        private string HexToString(string input)
        {
            StringBuilder sb = new StringBuilder();
            for (int i = 0; i < input.Length; i += 2)
            {
                string hs = input.Substring(i, 2);
                sb.Append(Convert.ToChar(Convert.ToUInt32(hs, 16)));
            }
            return sb.ToString();
        }

        private void cmbExchangeServerVersion_SelectedIndexChanged(object sender, EventArgs e)
        {
            bool enabled = false;
            switch (cmbExchangeServerVersion.SelectedItem.ToString())
            {
                case "2010":
                    enabled = true;
                    break;
                case "2013":
                    enabled = true;
                    break;
                case "Online":
                    txtExchangeServer.Text = string.Empty;
                    txtADDomain.Text = string.Empty;
                    enabled = false;
                    break;
            }
            txtExchangeServer.Enabled = enabled;
            txtADDomain.Enabled = enabled;
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            string fileName = AppConfigFileName.Substring(0, AppConfigFileName.LastIndexOf('\\') + 1) + "SavedConfiguration.xml";

            XmlDocument doc = new XmlDocument();
            doc.Load(fileName);

            XmlNodeList nodes = doc.SelectSingleNode("/Configuration/UserInfo").ChildNodes;

            foreach (XmlNode node in nodes)
            {
                switch (node.Name)
                {
                    case "Company":
                        node.InnerText = txtCompany.Text;
                        break;
                    case "WebSite":
                        node.InnerText = txtWebSiteUrl.Text;
                        break;
                    case "PhoneNumber":
                        node.InnerText = txtPhoneNumber.Text;
                        break;
                    case "FirstName":
                        node.InnerText = txtFirstName.Text;
                        break;
                    case "LastName":
                        node.InnerText = txtLastName.Text;
                        break;
                    case "EmailAddress":
                        node.InnerText = txtEmailAddress.Text;
                        break;
                }
            }

            nodes = doc.SelectSingleNode("/Configuration/CRMInfo").ChildNodes;
            foreach (XmlNode node in nodes)
            {
                switch (node.Name)
                {
                    case "AuthenticationType":
                        node.Attributes["Index"].Value = (cmbAuthenticationType.SelectedIndex + 1).ToString();
                        break;
                    case "CRMServiceUrl":
                        node.InnerText = txtCrmServiceUrl.Text;
                        break;
                    case "IntegratedAuthentication":
                        node.InnerText = chkIntegrated.Checked ? "Yes" : "No";
                        break;
                    case "CRMCredentials":
                        XmlNode userNameNode = node.FirstChild;
                        userNameNode.InnerText= txtCRMUsername.Text;

                        XmlNode domainNode = userNameNode.NextSibling;
                        domainNode.InnerText = txtCRMDomain.Text;
                        break;
                }
            }

            nodes = doc.SelectSingleNode("/Configuration/ExchangeInfo").ChildNodes;
            foreach (XmlNode node in nodes)
            {
                switch (node.Name)
                {
                    case "ExchangeServerUrl":
                        node.InnerText = txtExchangeServer.Text;
                        break;
                    case "ExchangeServerVersion":
                        node.InnerText = cmbExchangeServerVersion.SelectedItem.ToString();
                        break;
                    case "ExchangeCredentials":
                        XmlNode userNameNode = node.FirstChild;
                        userNameNode.InnerText = txtADUsername.Text;

                        XmlNode domainNode = userNameNode.NextSibling;
                        domainNode.InnerText = txtADDomain.Text;
                        break;
                    case "ExchangeDistributionGroupOU":
                        node.InnerText = cmbDistributionGroupOU.SelectedItem != null ? cmbDistributionGroupOU.SelectedItem.ToString() : string.Empty;
                        break;
                    case "ExchangeContactOU":
                        node.InnerText = cmbContactsOU.SelectedItem != null ? cmbContactsOU.SelectedItem.ToString() : string.Empty;
                        break;
                    case "InternetDomainSuffix":
                        node.InnerText = txtExternalDomainName.Text;
                        break;
                }
            }


            doc.Save(fileName);
        }

        private void btnLoadConfiguration_Click(object sender, EventArgs e)
        {
            string fileName = AppConfigFileName.Substring(0, AppConfigFileName.LastIndexOf('\\') + 1) + "SavedConfiguration.xml";

            XmlDocument doc = new XmlDocument();
            doc.Load(fileName);

            XmlNodeList nodes = doc.SelectSingleNode("/Configuration/UserInfo").ChildNodes;

            foreach (XmlNode node in nodes)
            {
                switch (node.Name)
                {
                    case "Company":
                        txtCompany.Text = node.InnerText;
                        break;
                    case "WebSite":
                        txtWebSiteUrl.Text = node.InnerText;
                        break;
                    case "PhoneNumber":
                        txtPhoneNumber.Text = node.InnerText;
                        break;
                    case "FirstName":
                        txtFirstName.Text = node.InnerText;
                        break;
                    case "LastName":
                        txtLastName.Text = node.InnerText;
                        break;
                    case "EmailAddress":
                        txtEmailAddress.Text = node.InnerText;
                        break;
                }
            }

            nodes = doc.SelectSingleNode("/Configuration/CRMInfo").ChildNodes;
            foreach (XmlNode node in nodes)
            {
                switch (node.Name)
                {
                    case "AuthenticationType":
                        cmbAuthenticationType.SelectedIndex = Convert.ToInt32(node.Attributes["Index"].Value) - 1;
                        break;
                    case "CRMServiceUrl":
                        txtCrmServiceUrl.Text = node.InnerText;
                        break;
                    case "IntegratedAuthentication":
                        if (node.InnerText == "Yes")
                            chkIntegrated.Checked = true;
                        break;
                    case "CRMCredentials":
                        XmlNode userNameNode = node.FirstChild;
                        txtCRMUsername.Text = userNameNode.InnerText;

                        XmlNode domainNode = userNameNode.NextSibling;
                        txtCRMDomain.Text = domainNode.InnerText;
                        break;
                }
            }

            nodes = doc.SelectSingleNode("/Configuration/ExchangeInfo").ChildNodes;
            foreach (XmlNode node in nodes)
            {
                switch (node.Name)
                {
                    case "ExchangeServerUrl":
                        txtExchangeServer.Text = node.InnerText;
                        break;
                    case "ExchangeServerVersion":
                        cmbExchangeServerVersion.SelectedItem = node.InnerText;
                        break;
                    case "ExchangeCredentials":
                        XmlNode userNameNode = node.FirstChild;
                        txtADUsername.Text = userNameNode.InnerText;

                        XmlNode domainNode = userNameNode.NextSibling;
                        txtADDomain.Text = domainNode.InnerText;
                        break;
                    case "ExchangeDistributionGroupOU":
                        lblDistributionGroupOU.Text = node.InnerText;
                        break;
                    case "ExchangeContactOU":
                        lblContactOU.Text = node.InnerText;
                        break;
                    case "InternetDomainSuffix":
                        txtExternalDomainName.Text = node.InnerText;
                        break;
                }
            }
        }

        private void btnExportLog_Click(object sender, EventArgs e)
        {
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
                excel.SaveConfigBook(book);
            }
        }

        private void ConfiguratorWizard_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (syncCore.crm != null)
                syncCore.crm.Disconnect();

            if (syncCore.exch != null)
                syncCore.exch.Disconnect();

            Application.Exit();
        }

        private void chkUseConnectionString_CheckedChanged(object sender, EventArgs e)
        {
            if (chkUseConnectionString.Checked)
            {
                int authenticationType = cmbAuthenticationType.SelectedIndex + 1;
                StringBuilder sb = new StringBuilder();
                switch (authenticationType)
                {
                    case 2:
                        if (ValidateCRMConnectionFields(authenticationType))
                        {
                            sb.Append(string.Format("Url=https://{0}/{1}; ", txtCrmServiceUrl.Text, txtOrganizationName.Text));
                            sb.Append(string.Format("Username={0}; ", txtCRMUsername.Text));
                            sb.Append(string.Format("Password={0}; ", "********"));
                            sb.Append(string.Format("AuthType=IFD"));
                            txtConnectionString.Text = sb.ToString();
                            sb.Clear();
                        }
                        break;
                    case 3:
                        sb.Append(string.Format("Url=https://{0}/{1}; ", txtCrmServiceUrl.Text, txtOrganizationName.Text));
                        sb.Append(string.Format("ClientId={0}; ", txtClientId.Text));
                        sb.Append(string.Format("ClientSecret={0}; ", txtClientSecret.Text));
                        sb.Append(string.Format("AuthType=ClientSecret"));
                        txtConnectionString.Text = sb.ToString();
                        break;
                    case 4:
                        if (ValidateCRMConnectionFields(authenticationType))
                        {
                            sb.Append(string.Format("Url=https://{0}; ", txtCrmServiceUrl.Text));
                            sb.Append(string.Format("Username={0}; ", txtCRMUsername.Text));
                            sb.Append(string.Format("Password={0}; ", "********"));
                            sb.Append(string.Format("AuthType=Office365"));
                            txtConnectionString.Text = sb.ToString();
                            sb.Clear();
                        }
                        break;
                    default:
                        MessageBox.Show("Connection Strings currently only available for Office 365 Connections", "Connection String Type Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        chkUseConnectionString.Checked = false;
                        break;

                }
            }
        }

        private void SetClientSecretDisplayMode(bool isVisible = true)
        {
            lblClientId.Visible = isVisible;
            lblClientSecret.Visible = isVisible;
            txtClientId.Visible = isVisible;
            txtClientSecret.Visible = isVisible;

            lblUsername.Visible = !isVisible;
            lblPassword.Visible = !isVisible;
            lblCRMDomain.Visible = !isVisible;
            txtCRMUsername.Visible = !isVisible;
            txtCRMPassword.Visible = !isVisible;
            txtCRMDomain.Visible = !isVisible;
        }

        private void btnChangePassword_Click(object sender, EventArgs e)
        {
            ChangePassword frm = new ChangePassword();
            frm.ShowDialog(this);
            
        }
    }

}
