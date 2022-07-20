using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;


namespace ExchangeSync
{
    public partial class CRMSettings : UserControl
    {
        public CRMSettings()
        {
            InitializeComponent();
        }

        private void btnSaveSettings_Click(object sender, EventArgs e)
        {
            SaveAppSettings();
            SaveConfigSettings();
            lblConfirmation.Text = "CRM Settings have been saved.";
        }

        private void CRMSettings_Load(object sender, EventArgs e)
        {
            LoadAppSettings();
            LoadConfigSettings();
        }

        private void LoadConfigSettings()
        {
            txtConnectionUrl.Text = ConfigSetting.CRMServiceUrl;

            txtCRMUsername.Text = ConfigSetting.CRMCredentials.UserName; 
            txtCRMDomain.Text = ConfigSetting.CRMCredentials.Domain;

            AuthenticationProviderType authenticationProviderType = ConfigSetting.AuthenticationProvider;  
            if (authenticationProviderType.ToInt() != 0)
                cmbAuthenticationType.SelectedIndex = authenticationProviderType.ToInt() - 1;

            bool integratedSecurity = ConfigSetting.IntegratedAuthentication; 
            chkIntegrated.Checked = integratedSecurity;
        }

        private void LoadAppSettings()
        {
            txtRequiredAttributes.Text = AppSetting.ContactRequiredAttributes.Value;
            txtCustomField.Text = AppSetting.ContactAutoNumberFieldName.Value;
            txtExchangeAliasField.Text = AppSetting.ContactExchangeAliasFieldName.Value;
            txtRevisionField.Text = AppSetting.ContactRevisionFieldName.Value;
        }

        private void SaveConfigSettings()
        {
            Configuration config = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);

            // Update CRM Fields
            KeyValueConfigurationCollection settings = config.AppSettings.Settings;
            settings["CRMServiceUrl"].Value = txtConnectionUrl.Text;
            ConfigSetting.CRMServiceUrl = txtConnectionUrl.Text;

            settings["CRMUsername"].Value = txtCRMUsername.Text;
            ConfigSetting.CRMCredentials.UserName = txtCRMUsername.Text;

            settings["CRMDomain"].Value = txtCRMDomain.Text;
            ConfigSetting.CRMCredentials.Domain = txtCRMDomain.Text;

            settings["CRMAuthenticationProviderType"].Value = (cmbAuthenticationType.SelectedIndex + 1).ToString();
            ConfigSetting.AuthenticationProvider = (AuthenticationProviderType)(cmbAuthenticationType.SelectedIndex + 1);

            settings["CRMIntegratedAuthentication"].Value = chkIntegrated.Checked ? "1" : "0";
            ConfigSetting.IntegratedAuthentication = chkIntegrated.Checked ? true : false;

            config.Save(ConfigurationSaveMode.Modified);
            ConfigurationManager.RefreshSection("appSettings");

            /*
            ConfigurationSection section = config.GetSection(Program.ENCRYPTED_SECTION_NAME);
            ClientSettingsSection clientSection = (ClientSettingsSection)section;

            SettingElement setPassword = clientSection.Settings.Get("CRMPassword");
            setPassword.Value.ValueXml.InnerXml = txtCRMPassword.Text;

            section.SectionInformation.ForceSave = true;
            config.Save(ConfigurationSaveMode.Modified);
            ConfigurationManager.RefreshSection(Program.ENCRYPTED_SECTION_NAME);
            */
        }

        private void SaveAppSettings()
        {
            Core syncCore = ExchangeSyncManager.SyncCore;
            // Required Attributes Field
            Guid key = AppSetting.ContactRequiredAttributes.Key;
            AppSetting.ContactRequiredAttributes = new KeyValuePair<Guid, string>(key, txtRequiredAttributes.Text);
            syncCore.crm.UpdateApplicationSetting(key, txtRequiredAttributes.Text);

            // Auto Number / Custom Field
            key = AppSetting.ContactAutoNumberFieldName.Key;
            AppSetting.ContactAutoNumberFieldName = new KeyValuePair<Guid, string>(key, txtCustomField.Text);
            syncCore.crm.UpdateApplicationSetting(key, txtCustomField.Text);

            // Revision Field
            key = AppSetting.ContactRevisionFieldName.Key;
            AppSetting.ContactRevisionFieldName = new KeyValuePair<Guid, string>(key, txtRevisionField.Text);
            syncCore.crm.UpdateApplicationSetting(key, txtRevisionField.Text);

            // Revision Field
            key = AppSetting.ContactExchangeAliasFieldName.Key;
            AppSetting.ContactExchangeAliasFieldName = new KeyValuePair<Guid, string>(key, txtExchangeAliasField.Text);
            syncCore.crm.UpdateApplicationSetting(key, txtExchangeAliasField.Text);
        }

        private void cmbAuthenticationType_SelectedIndexChanged(object sender, EventArgs e)
        {
            switch (cmbAuthenticationType.SelectedIndex)
            {
                case 0: // On-Premise 
                    chkIntegrated.Checked = true;
                    txtCRMDomain.ReadOnly = true; txtCRMDomain.Enabled = false;
                    txtCRMUsername.ReadOnly = true; txtCRMUsername.Enabled = false;
                    //txtCRMPassword.ReadOnly = true; txtCRMPassword.Enabled = false;
                    break;
                case 1: // Federated (On-Premise, Partner Hosted or CRM Online)
                    chkIntegrated.Checked = false;
                    txtCRMDomain.Text = ""; txtCRMDomain.ReadOnly = true; txtCRMDomain.Enabled = false;
                    break;
                case 2: // CRM Online Legacy
                    chkIntegrated.Checked = false;
                    txtCRMDomain.Text = ""; txtCRMDomain.ReadOnly = true; txtCRMDomain.Enabled = false;
                    break;
            }
        }

        private void chkIntegrated_CheckedChanged(object sender, EventArgs e)
        {
            if (chkIntegrated.Checked)
            {
                txtCRMUsername.Text = ""; txtCRMUsername.ReadOnly = true; txtCRMUsername.Enabled = false;
                //txtCRMPassword.Text = ""; txtCRMPassword.ReadOnly = true; txtCRMPassword.Enabled = false;
                txtCRMDomain.Text = ""; txtCRMDomain.ReadOnly = true; txtCRMDomain.Enabled = false;
            }
            else
            {
                txtCRMUsername.ReadOnly = false; txtCRMUsername.Enabled = true;
                //txtCRMPassword.ReadOnly = false; txtCRMPassword.Enabled = true;
                txtCRMDomain.ReadOnly = false; txtCRMDomain.Enabled = true;
            }
        }

        private void btnUpdateSolution_Click(object sender, EventArgs e)
        {
            OpenFileDialog dlg = new OpenFileDialog();
            dlg.InitialDirectory = System.IO.Path.GetDirectoryName(Application.ExecutablePath);
            dlg.Filter = "Zip Files (*.zip)|*.zip";
            dlg.Multiselect = false;

            DialogResult result = dlg.ShowDialog();

            if (result == DialogResult.OK)
            {
                string fileName = dlg.FileName;
                ExchangeSyncManager.SyncCore.crm.ImportSolution(fileName);
            }
        }

        private void btnChangePassword_Click(object sender, EventArgs e)
        {
            //txtCRMPassword.Enabled = true;
            //txtCRMPassword.ReadOnly = false;
            //txtCRMPassword.Text = "";
        }
    }
}
