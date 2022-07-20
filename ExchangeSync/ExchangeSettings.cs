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
    public partial class ExchangeSettings : UserControl
    {
        public ExchangeSettings()
        {
            InitializeComponent();
        }
        private bool PasswordChanged = false;
        private void ExchangeSettings_Load(object sender, EventArgs e)
        {
            LoadAppSettings();
            LoadConfigSettings();
        }

        private void LoadConfigSettings()
        {
            ExchangeServerType version = ConfigSetting.ExchangeServerVersion;
            if (version != ExchangeServerType.Undefined)
            {
                switch (version)
                {
                    case ExchangeServerType.Exchange_2010:
                        cmbVersion.Text = "Exchange 2010";
                        cmbVersion.SelectedIndex = 0;
                        break;
                    case ExchangeServerType.Exchange_2013:
                        cmbVersion.Text = "Exchange 2013";
                        cmbVersion.SelectedIndex = 1;
                        break;
                    case ExchangeServerType.Exchange_2016:
                        cmbVersion.Text = "Exchange 2016";
                        cmbVersion.SelectedIndex = 2;
                        break;
                    case ExchangeServerType.Exchange_Online_365:
                        cmbVersion.Text = "Office 365";
                        cmbVersion.SelectedIndex = 3;
                        break;

                }
            }

            if (cmbVersion.Text != "Office 365")
            {
                txtServer.Text = ConfigSetting.ExchangeServerUrl;
                txtUsername.Text = ConfigSetting.ExchangeCredentials.UserName;
                txtDomain.Text = ConfigSetting.ExchangeCredentials.Domain;
            }
            else
            {
                txtServer.Enabled = false;
                txtUsername.Text = ConfigSetting.ExchangeCredentials.UserName;
                txtDomain.Enabled = false;
            }
        }

        private void LoadAppSettings()
        {
            if (AppSetting.ExchangeAliasFieldFormat.Value != "APPENDDOMAIN")
                cmbAliasType.Text = AppSetting.ExchangeAliasFieldFormat.Value;
            else
                cmbAliasType.Text = "Append Domain";

            if (AppSetting.ExchangeNameFieldFormat.Value != "APPENDDOMAIN")
                cmbNameType.Text = AppSetting.ExchangeAliasFieldFormat.Value;
            else
                cmbNameType.Text = "Append Domain";

            txtAliasValue.Text = AppSetting.ExchangeAliasFieldValue.Value;
            txtNameValue.Text = AppSetting.ExchangeNameFieldValue.Value;
            cmbAutoUpdatePolicy.Text = AppSetting.ExchangePolicyAutoUpdateEmailAddresses.Value.ToPascalCase();

            if (cmbVersion.Text != "Online")
            {
                string dlOU = AppSetting.ExchangeDistributionGroupOU.Value;
                string contactsOU = AppSetting.ExchangeMailContactOU.Value;

                LoadExchangeOrgUnits();

                for (int i = 0; i < uceDistributionGroups.Items.Count; i++)
                {
                    if (uceDistributionGroups.Items[i].DataValue.ToString() == dlOU)
                        uceDistributionGroups.SelectedIndex = i;
                }
                for (int i = 0; i < uceContacts.Items.Count; i++)
                {
                    if (uceContacts.Items[i].DataValue.ToString() == contactsOU)
                        uceContacts.SelectedIndex = i;
                }
            }
            else
            {
                uceDistributionGroups.Enabled = false;
                uceContacts.Enabled = false;
            }
            


            if (AppSetting.InternalRecipientType.Value.ToUpper() == "MAILBOX-ENABLED")
                rbMailboxEnabled.Checked = true;
            else if (AppSetting.InternalRecipientType.Value.ToUpper() == "MAIL-ENABLED")
                rbMailEnabled.Checked = true;
            else
                rbMailboxEnabled.Checked = true; // default


            LoadConfigSettings();
        }

        private void LoadExchangeOrgUnits()
        {

            List<OrganizationalUnit> list = ExchangeSyncManager.SyncCore.exch.GetAllOrganizationalUnits();
            foreach (OrganizationalUnit unit in list)
            {
                uceDistributionGroups.Items.Add(unit.CanonicalName);
                uceContacts.Items.Add(unit.CanonicalName);
            }
        }

        private void btnSaveSettings_Click(object sender, EventArgs e)
        {
            if (ValidateSettings())
            {
                SaveAppSettings();
                SaveConfigSettings();
                lblConfirmation.Text = "Exchange Settings have been saved.";
            }
            else
            {
                MessageBox.Show("Errors encountered during validation of your system settings. Save Operation Cancelled!", "Save Errors", MessageBoxButtons.OK);
            }
        }

        private void SaveConfigSettings()
        {
            Configuration config = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);

            KeyValueConfigurationCollection settings = config.AppSettings.Settings;
            settings["ADUsername"].Value = txtUsername.Text;
            ConfigSetting.ExchangeCredentials.UserName = txtUsername.Text;

            settings["ADDomain"].Value = txtDomain.Text;
            ConfigSetting.ExchangeCredentials.Domain = txtDomain.Text;

            settings["ExchangeServerName"].Value = txtServer.Text;
            ConfigSetting.ExchangeServerUrl = txtServer.Text;

            settings["ExchangeServerVersion"].Value = cmbVersion.Text;
            ConfigSetting.ExchangeServerVersion = (ExchangeServerType)cmbVersion.SelectedValue;


            config.Save(ConfigurationSaveMode.Modified);
            ConfigurationManager.RefreshSection("appSettings");
        }

        private void SaveAppSettings()
        {
            Core syncCore = ExchangeSyncManager.SyncCore;

            Guid key = Guid.Empty;

            if (cmbVersion.Text != "Online")
            {
                // Distribution Group OU
                key = AppSetting.ExchangeDistributionGroupOU.Key;
                AppSetting.ExchangeDistributionGroupOU = new KeyValuePair<Guid, string>(key, uceDistributionGroups.SelectedItem.DataValue.ToString());
                syncCore.crm.UpdateApplicationSetting(key, uceContacts.SelectedItem.DataValue.ToString());

                // Mail Contacts OU
                key = AppSetting.ExchangeMailContactOU.Key;
                AppSetting.ExchangeMailContactOU = new KeyValuePair<Guid, string>(key, uceContacts.SelectedItem.DataValue.ToString());
                syncCore.crm.UpdateApplicationSetting(key, uceContacts.SelectedItem.DataValue.ToString());
            }

            // Exchange Name Field Format/Type
            key = AppSetting.ExchangeNameFieldFormat.Key;
            AppSetting.ExchangeNameFieldFormat = new KeyValuePair<Guid, string>(key, cmbNameType.Text.StripPunctuation().ToUpper());
            syncCore.crm.UpdateApplicationSetting(key, cmbNameType.Text.StripPunctuation().ToUpper());

            // Exchange Name Field Value
            key = AppSetting.ExchangeNameFieldValue.Key;
            AppSetting.ExchangeNameFieldValue = new KeyValuePair<Guid, string>(key, txtNameValue.Text);
            syncCore.crm.UpdateApplicationSetting(key, txtNameValue.Text);

            // Exchange Alias Field Format/Type
            key = AppSetting.ExchangeAliasFieldFormat.Key;
            AppSetting.ExchangeAliasFieldFormat = new KeyValuePair<Guid, string>(key, cmbAliasType.Text.StripPunctuation().ToUpper());
            syncCore.crm.UpdateApplicationSetting(key, cmbAliasType.Text.StripPunctuation().ToUpper());

            // Exchange Alias Field Value
            key = AppSetting.ExchangeAliasFieldValue.Key;
            AppSetting.ExchangeAliasFieldValue = new KeyValuePair<Guid, string>(key, txtAliasValue.Text);
            syncCore.crm.UpdateApplicationSetting(key, txtAliasValue.Text);

            // Exchange Auto Update Email Addresses Based on Policy
            key = AppSetting.ExchangePolicyAutoUpdateEmailAddresses.Key;
            AppSetting.ExchangePolicyAutoUpdateEmailAddresses = new KeyValuePair<Guid, string>(key, cmbAutoUpdatePolicy.Text.ToUpper());
            syncCore.crm.UpdateApplicationSetting(key, cmbAutoUpdatePolicy.Text.ToUpper());

            key = AppSetting.InternalRecipientType.Key;
            if (rbMailboxEnabled.Checked)
            {
                AppSetting.InternalRecipientType = new KeyValuePair<Guid, string>(key, "Mailbox-Enabled");
                syncCore.crm.UpdateApplicationSetting(key, "Mailbox-Enabled");
            }
            else if (rbMailEnabled.Checked)
            {
                AppSetting.InternalRecipientType = new KeyValuePair<Guid, string>(key, "Mail-Enabled");
                syncCore.crm.UpdateApplicationSetting(key, "Mail-Enabled");
            }
            
            SaveConfigSettings();
        }

        private bool ValidateSettings()
        {
            bool rc = true;

            if (string.IsNullOrEmpty(txtUsername.Text))
                rc = false;

            if (PasswordChanged)
            {
                //if (string.IsNullOrEmpty(txtPassword.Text))
                    //rc = false;
            }

            if (string.IsNullOrEmpty(cmbVersion.Text))
                rc = false;

            if (cmbVersion.Text != "Online")
            {
                if (string.IsNullOrEmpty(txtServer.Text))
                    rc = false;
                
                if (string.IsNullOrEmpty(uceDistributionGroups.SelectedItem.DataValue.ToString()))
                    rc = false;

                if (string.IsNullOrEmpty(uceContacts.SelectedItem.DataValue.ToString()))
                    rc = false;
            }
            return rc;
        }

        private void btnChangePassword_Click(object sender, EventArgs e)
        {


        }

        private void ExportExchangeGroups()
        {
            Core syncCore = ExchangeSyncManager.SyncCore;

            if (!Trace.Null)
                Trace.ClearLog();
            else
                Trace.InitializeLog();

            Trace.AddLog("\"Group Name\",\"Contact Name\",\"Email Address\",\"Alias\"");
            List<DistributionGroup> groups = syncCore.exch.GetAllDistributionListGroups();
            foreach (DistributionGroup group in groups)
            {
                string groupName = group.Name;
                List<ExchangeContact> contacts = syncCore.exch.GetDistributionListMembers(groupName);
                foreach (ExchangeContact contact in contacts)
                {
                    string name = contact.Name;
                    string emailAddress = contact.PrimarySmtpAddress;
                    string alias = contact.Alias;
                    string identity = contact.Identity;
                    string ou = contact.OrganizationalUnit;
                    Trace.AddLog(String.Format("\"{0}\",\"{1}\"\"{2}\"\"{3}\"", groupName, name, emailAddress, alias));
                }
            }

            // Check if logs directory exists
            if (!System.IO.Directory.Exists(@"c:\logs"))
            {
                System.IO.Directory.CreateDirectory(@"C:\logs");
            }

            string filename = @"C:\logs\exchangegroups_" + DateTime.Now.ToString("yyMMddHHmmss") + ".txt";
            System.IO.File.WriteAllText(filename, Trace.RetrieveLogAsText(false));

            Trace.ClearLog();
        }

        private void btnExportExchangeGroups_Click(object sender, EventArgs e)
        {
            ExportExchangeGroups();
        }

        private void btnChangePassword_Click_1(object sender, EventArgs e)
        {
            Password frm = new Password();
            frm.ShowDialog(this);
        }

        private void cmbNameType_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

    }
}
