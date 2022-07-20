using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace ExchangeSync
{
    public partial class Options : Form
    {
        Core syncCore = ExchangeSyncManager.SyncCore;

        public Options()
        {
            InitializeComponent();
        }

        private void label16_Click(object sender, EventArgs e)
        {

        }

        private void btnOK_Click(object sender, EventArgs e)
        {
            SaveOptions();

        }

        private void ebOptions_GroupClick(object sender, Infragistics.Win.UltraWinExplorerBar.GroupEventArgs e)
        {
            switch (e.Group.Key)
            {
                case "General":
                    utOptions.SelectedTab = utOptions.Tabs["tabGeneral"];
                    break;
                case "CRM":
                    utOptions.SelectedTab = utOptions.Tabs["tabCRM"];
                    break;
                case "Exchange":
                    utOptions.SelectedTab = utOptions.Tabs["tabExchange"];
                    break;
                case "Filtering":
                    utOptions.SelectedTab = utOptions.Tabs["tabMappingFilter"];
                    break;
                case "UsersContacts":
                    utOptions.SelectedTab = utOptions.Tabs["tabUsersContacts"];
                    break;
                case "System":
                    utOptions.SelectedTab = utOptions.Tabs["tabSettings"];
                    break;
                case "Trace":
                    utOptions.SelectedTab = utOptions.Tabs["tabTrace"];
                    break;

            }
        }

        private void label31_Click(object sender, EventArgs e)
        {

        }

        private void Options_Load(object sender, EventArgs e)
        {
            LoadSettings();
        }

        private void LoadSettings()
        {
            // General Settings
            txtListDisplayName.Text = AppSetting.ListDisplayName.GetValue();
            
            // CRM Settings
            txtConnectionUrl.Text = ConfigSetting.CRMServiceUrl;
            cmbAuthenticationType.SelectedIndex = (ConfigSetting.AuthenticationProvider.ToInt() - 1);
            txtRequiredAttributes.Text = AppSetting.ContactRequiredAttributes.GetValue();
            txtCustomField.Text = AppSetting.ContactAutoNumberFieldName.GetValue();
            txtExchangeAliasField.Text = AppSetting.ContactExchangeAliasFieldName.GetValue();
            txtRevisionField.Text = AppSetting.ContactRevisionFieldName.GetValue();

            // Exchange Settings
            txtServer.Text = ConfigSetting.ExchangeServerUrl;
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

                cmbAutoUpdatePolicy.Text = AppSetting.ExchangePolicyAutoUpdateEmailAddresses.GetValue().ToPascalCase();
                cmbHideFromAddressBook.Text = AppSetting.HideFromExchangeAddressLists.GetValue().ToPascalCase();

                if (version != ExchangeServerType.Exchange_Online_365)
                {
                    string dlOU = AppSetting.ExchangeDistributionGroupOU.GetValue();
                    string contactsOU = AppSetting.ExchangeMailContactOU.GetValue();

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

                if (AppSetting.InternalRecipientType.GetValue().ToUpper() == "MAILBOX-ENABLED")
                    rbMailboxEnabled.Checked = true;
                else if (AppSetting.InternalRecipientType.GetValue().ToUpper() == "MAIL-ENABLED")
                    rbMailEnabled.Checked = true;
                else
                    rbMailboxEnabled.Checked = true; // default

                if (AppSetting.ExchangeMailUserUpdate.GetValue().ToUpper() == "YES")
                    chkUpdateMailUser.Checked = true;
                else
                    chkUpdateMailUser.Checked = false;

                // Field Mapping
                switch (AppSetting.FieldMappingAction.GetValue().ToUpper())
                {
                    case "NONE":
                        rbMapNone.Checked = true;
                        rbMapMinimal.Checked = false;
                        rbMapCRM.Checked = false;
                        break;
                    case "MIN":
                        rbMapNone.Checked = false;
                        rbMapMinimal.Checked = true;
                        rbMapCRM.Checked = false;
                        break;
                    case "CRM":
                        rbMapNone.Checked = false;
                        rbMapMinimal.Checked = false;
                        rbMapCRM.Checked = true;
                        break;
                }

                switch (AppSetting.ExchangeUpdateAction.GetValue().ToUpper())
                {
                    case "DELTA":
                        rbUpdateDelta.Checked = true;
                        rbUpdateAll.Checked = false;
                        break;
                    case "ALL":
                        rbUpdateDelta.Checked = false;
                        rbUpdateAll.Checked = true;
                        break;
                }

                // Add AppSetting Exchange Identity Type (Alias, Email Address or Exchange GUID)
                // AppSetting.ExchangeIdentityType.Value

                // Users and Contacts
                if (AppSetting.ExchangeAliasFieldFormat.GetValue() != "APPENDDOMAIN")
                    cmbAliasType.Text = AppSetting.ExchangeAliasFieldFormat.GetValue();
                else
                    cmbAliasType.Text = "Append Domain";

                if (AppSetting.ExchangeNameFieldFormat.GetValue() != "APPENDDOMAIN")
                    cmbNameType.Text = AppSetting.ExchangeAliasFieldFormat.GetValue();
                else
                    cmbNameType.Text = "Append Domain";

                txtAliasValue.Text = AppSetting.ExchangeAliasFieldValue.GetValue();
                txtNameValue.Text = AppSetting.ExchangeNameFieldValue.GetValue();

                switch (AppSetting.ExchangeMailContactFormat.GetValue().ToUpper())
                {
                    case "NOUPDATE":
                        cmbMailContactFormat.SelectedIndex = 0;
                        // cmbMailContactFormat.SelectedText = "No Update";
                        break;
                    case "CRMSETTING":
                        cmbMailContactFormat.SelectedIndex = 1;
                        // cmbMailContactFormat.SelectedText = "CRM Setting";
                        break;
                    case "FIRSTLAST":
                        cmbMailContactFormat.SelectedIndex = 2;
                        // cmbMailContactFormat.SelectedText = "First Last";
                        break;
                    case "LASTFIRST":
                        cmbMailContactFormat.SelectedIndex = 3;
                        // cmbMailContactFormat.SelectedText = "Last, First";
                        break;
                }

                switch (AppSetting.ExchangeMailUserFormat.GetValue().ToUpper())
                {
                    case "NOUPDATE":
                        cmbMailUserFormat.SelectedIndex = 0;
                        // cmbMailUserFormat.SelectedText = "No Update";
                        break;
                    case "CRMSETTING":
                        cmbMailUserFormat.SelectedIndex = 1;
                        // cmbMailUserFormat.SelectedText = "CRM Settings";
                        break;
                    case "FIRSTLAST":
                        cmbMailUserFormat.SelectedIndex = 2;
                        // cmbMailUserFormat.SelectedText = "First Last";
                        break;
                    case "LASTFIRST":
                        cmbMailUserFormat.SelectedIndex = 3;
                        // cmbMailUserFormat.SelectedText = "Last, First";
                        break;
                }

                if (AppSetting.InternalRecipientType.GetValue().ToUpper() == "MAILBOX-ENABLED")
                {
                    rbMailboxEnabled.Checked = true;
                    rbMailEnabled.Checked = false;
                }
                else if (AppSetting.InternalRecipientType.GetValue().ToUpper() == "MAIL-ENABLED")
                {
                    rbMailboxEnabled.Checked = false;
                    rbMailEnabled.Checked = true;
                }

                if (AppSetting.ExchangeIdentityType.GetValue().ToUpper() == "ALIAS")
                {
                    rbExchangeAlias.Checked = true;
                    rbEmailAddress.Checked = false;
                }
                else if (AppSetting.ExchangeIdentityType.GetValue().ToUpper() == "EMAILADDRESS")
                {
                    rbExchangeAlias.Checked = false;
                    rbEmailAddress.Checked = true;
                }


                if (!string.IsNullOrEmpty(AppSetting.LoggingCriteria.GetValue()))
                {
                    char[] criteria = AppSetting.LoggingCriteria.GetValue().ToCharArray();
                    if (criteria.Length == 4)
                    {
                        chkInformation.Checked = criteria[0] == '1' ? true : false;
                        chkWarning.Checked = criteria[1] == '1' ? true : false;
                        chkError.Checked = criteria[2] == '1' ? true : false;
                        chkCritical.Checked = criteria[3] == '1' ? true : false;
                    }
                }

                if (!string.IsNullOrEmpty(AppSetting.LoggingCriteriaVerbose.GetValue()))
                {
                    string value = AppSetting.LoggingCriteriaVerbose.GetValue().ToLower();
                    if (value == "yes")
                        chkVerbose.Checked = true;
                    else
                        chkVerbose.Checked = false;
                }

                if (!string.IsNullOrEmpty(AppSetting.RestartAutoSyncServicePostOperation.GetValue()))
                {
                    string value = AppSetting.RestartAutoSyncServicePostOperation.GetValue().ToLower();
                    if (value == "yes")
                        chkRestartService.Checked = true;
                    else
                        chkRestartService.Checked = false;
                }

                if (!string.IsNullOrEmpty(AppSetting.DeleteChangeTracking.GetValue()))
                {
                    string value = AppSetting.DeleteChangeTracking.GetValue().ToLower();
                    if (value == "yes")
                        chkDeleteChangeTracking.Checked = true;
                    else
                        chkDeleteChangeTracking.Checked = false;
                }

                if (!string.IsNullOrEmpty(AppSetting.DeleteExchangeContacts.GetValue()))
                {
                    string value = AppSetting.DeleteExchangeContacts.GetValue().ToLower();
                    if (value == "yes")
                        chkDeleteOrphanedContacts.Checked = true;
                    else
                        chkDeleteOrphanedContacts.Checked = false;
                }



            }
        }

        private void SaveOptions()
        {
            // General Options Tab
            if (txtListDisplayName.Text != AppSetting.ListDisplayName.GetValue())
            {
                Guid appSettingsId = AppSetting.ListDisplayName.Key;
                AppSetting.ListDisplayName = new KeyValuePair<Guid, string>(appSettingsId, txtListDisplayName.Text);
                syncCore.crm.UpdateApplicationSetting(AppSetting.ListDisplayName.Key, AppSetting.ListDisplayName.GetValue());
            }

            // CRM Options Tab
            if (txtRequiredAttributes.Text != AppSetting.ContactRequiredAttributes.GetValue())
            {
                Guid appSettingsId = AppSetting.ContactRequiredAttributes.Key;
                AppSetting.ContactRequiredAttributes = new KeyValuePair<Guid, string>(appSettingsId, txtRequiredAttributes.Text);
                syncCore.crm.UpdateApplicationSetting(AppSetting.ContactRequiredAttributes.Key, AppSetting.ContactRequiredAttributes.GetValue());
            }

            if (txtCustomField.Text != AppSetting.ContactAutoNumberFieldName.GetValue())
            {
                Guid appSettingsId = AppSetting.ContactAutoNumberFieldName.Key;
                AppSetting.ContactAutoNumberFieldName = new KeyValuePair<Guid, string>(appSettingsId, txtCustomField.Text);
                syncCore.crm.UpdateApplicationSetting(AppSetting.ContactAutoNumberFieldName.Key, AppSetting.ContactAutoNumberFieldName.GetValue());
            }

            if (txtExchangeAliasField.Text != AppSetting.ContactExchangeAliasFieldName.GetValue())
            {
                Guid appSettingsId = AppSetting.ContactExchangeAliasFieldName.Key;
                AppSetting.ContactExchangeAliasFieldName = new KeyValuePair<Guid, string>(appSettingsId, txtExchangeAliasField.Text);
                syncCore.crm.UpdateApplicationSetting(AppSetting.ContactExchangeAliasFieldName.Key, AppSetting.ContactExchangeAliasFieldName.GetValue());
            }

            if (txtRevisionField.Text != AppSetting.ContactRevisionFieldName.GetValue())
            {
                Guid appSettingsId = AppSetting.ContactRevisionFieldName.Key;
                AppSetting.ContactRevisionFieldName = new KeyValuePair<Guid, string>(appSettingsId, txtRevisionField.Text);
                syncCore.crm.UpdateApplicationSetting(AppSetting.ContactRevisionFieldName.Key, AppSetting.ContactRevisionFieldName.GetValue());
            }

            // Exchange Options Tab
            if (uceDistributionGroups.Enabled)
            {
                if (uceDistributionGroups.SelectedItem.DataValue.ToString() != AppSetting.ExchangeDistributionGroupOU.GetValue())
                {
                    Guid appSettingId = AppSetting.ExchangeDistributionGroupOU.Key;
                    AppSetting.ExchangeDistributionGroupOU = new KeyValuePair<Guid, string>(appSettingId, uceDistributionGroups.SelectedItem.DataValue.ToString());
                    syncCore.crm.UpdateApplicationSetting(AppSetting.ExchangeDistributionGroupOU.Key, AppSetting.ExchangeDistributionGroupOU.GetValue());
                }
            }

            if (uceContacts.Enabled)
            {
                if (uceContacts.SelectedItem.DataValue.ToString() != AppSetting.ExchangeMailContactOU.GetValue())
                {
                    Guid appSettingId = AppSetting.ExchangeMailContactOU.Key;
                    AppSetting.ExchangeMailContactOU = new KeyValuePair<Guid, string>(appSettingId, uceContacts.SelectedItem.DataValue.ToString());
                    syncCore.crm.UpdateApplicationSetting(AppSetting.ExchangeMailContactOU.Key, AppSetting.ExchangeMailContactOU.GetValue());
                }
            }

            if (cmbAutoUpdatePolicy.Text.ToUpper() != AppSetting.ExchangeMailContactOU.GetValue())
            {
                Guid appSettingId = AppSetting.ExchangePolicyAutoUpdateEmailAddresses.Key;
                AppSetting.ExchangePolicyAutoUpdateEmailAddresses = new KeyValuePair<Guid, string>(appSettingId, cmbAutoUpdatePolicy.Text.ToUpper());
                syncCore.crm.UpdateApplicationSetting(AppSetting.ExchangePolicyAutoUpdateEmailAddresses.Key, AppSetting.ExchangePolicyAutoUpdateEmailAddresses.GetValue());
            }

            if (cmbHideFromAddressBook.Text.ToUpper() != AppSetting.HideFromExchangeAddressLists.GetValue())
            {
                Guid appSettingsId = AppSetting.HideFromExchangeAddressLists.Key;
                AppSetting.HideFromExchangeAddressLists = new KeyValuePair<Guid, string>(appSettingsId, cmbHideFromAddressBook.Text.ToUpper());
                syncCore.crm.UpdateApplicationSetting(AppSetting.HideFromExchangeAddressLists.Key, AppSetting.HideFromExchangeAddressLists.GetValue());
            }


                // Field Mapping and Sync Filter Tab
                string optionName = string.Empty;
            if (rbMapNone.Checked)
                optionName = "NONE";
            else if (rbMapMinimal.Checked)
                optionName = "MIN";
            else if (rbMapCRM.Checked)
                optionName = "CRM";

            if (optionName != AppSetting.FieldMappingAction.GetValue())
            {
                Guid appSettingsId = AppSetting.FieldMappingAction.Key;
                AppSetting.FieldMappingAction = new KeyValuePair<Guid, string>(appSettingsId, optionName);
                syncCore.crm.UpdateApplicationSetting(AppSetting.FieldMappingAction.Key, AppSetting.FieldMappingAction.GetValue());
            }

            optionName = string.Empty;
            if (rbUpdateAll.Checked)
                optionName = "ALL";
            else if (rbUpdateDelta.Checked)
                optionName = "DELTA";

            if (optionName != AppSetting.ExchangeUpdateAction.Value)
            {
                Guid appSettingsId = AppSetting.ExchangeUpdateAction.Key;
                AppSetting.ExchangeUpdateAction = new KeyValuePair<Guid, string>(appSettingsId, optionName);
                syncCore.crm.UpdateApplicationSetting(AppSetting.ExchangeUpdateAction.Key, AppSetting.ExchangeUpdateAction.Value);
            }

            // Users and Contacts
            if (!string.IsNullOrEmpty(cmbMailContactFormat.Text))
            {
                if (cmbMailContactFormat.Text != AppSetting.ExchangeMailContactFormat.GetValue())
                {
                    Guid appSettingsId = AppSetting.ExchangeMailContactFormat.Key;
                    string mailContactFormat = cmbMailContactFormat.Text.StripPunctuation().ToUpper();
                    AppSetting.ExchangeMailContactFormat = new KeyValuePair<Guid, string>(appSettingsId, mailContactFormat);
                    syncCore.crm.UpdateApplicationSetting(AppSetting.ExchangeMailContactFormat.Key, AppSetting.ExchangeMailContactFormat.GetValue());
                }
            }

            if (!string.IsNullOrEmpty(cmbMailContactFormat.Text))
            {
                if (cmbMailContactFormat.Text != AppSetting.ExchangeMailUserFormat.GetValue())
                {
                    Guid appSettingsId = AppSetting.ExchangeMailUserFormat.Key;
                    string mailUserFormat = cmbMailUserFormat.Text.StripPunctuation().ToUpper();
                    AppSetting.ExchangeMailUserFormat = new KeyValuePair<Guid, string>(appSettingsId, mailUserFormat);
                    syncCore.crm.UpdateApplicationSetting(AppSetting.ExchangeMailUserFormat.Key, AppSetting.ExchangeMailUserFormat.GetValue());
                }
            }

            if (cmbNameType.Text.StripPunctuation().ToUpper() != AppSetting.ExchangeNameFieldFormat.GetValue())
            {
                Guid appSettingsId = AppSetting.ExchangeNameFieldFormat.Key;
                AppSetting.ExchangeNameFieldFormat = new KeyValuePair<Guid, string>(appSettingsId, cmbNameType.Text.StripPunctuation().ToUpper());
                syncCore.crm.UpdateApplicationSetting(AppSetting.ExchangeNameFieldFormat.Key, AppSetting.ExchangeNameFieldFormat.GetValue());
            }

            if (txtNameValue.Text != AppSetting.ExchangeNameFieldValue.GetValue())
            {
                Guid appSettingsId = AppSetting.ExchangeNameFieldValue.Key;
                AppSetting.ExchangeNameFieldValue = new KeyValuePair<Guid, string>(appSettingsId, txtNameValue.Text);
                syncCore.crm.UpdateApplicationSetting(AppSetting.ExchangeNameFieldValue.Key, AppSetting.ExchangeNameFieldValue.GetValue());
            }

            if (cmbAliasType.Text.StripPunctuation().ToUpper() != AppSetting.ExchangeAliasFieldFormat.GetValue())
            {
                Guid appSettingsId = AppSetting.ExchangeAliasFieldFormat.Key;
                AppSetting.ExchangeAliasFieldFormat = new KeyValuePair<Guid, string>(appSettingsId, cmbAliasType.Text.StripPunctuation().ToUpper());
                syncCore.crm.UpdateApplicationSetting(AppSetting.ExchangeAliasFieldFormat.Key, AppSetting.ExchangeAliasFieldFormat.GetValue());
            }

            if (txtAliasValue.Text != AppSetting.ExchangeAliasFieldValue.GetValue())
            {
                Guid appSettingsId = AppSetting.ExchangeAliasFieldValue.Key;
                AppSetting.ExchangeAliasFieldValue = new KeyValuePair<Guid, string>(appSettingsId, txtAliasValue.Text);
                syncCore.crm.UpdateApplicationSetting(AppSetting.ExchangeAliasFieldValue.Key, AppSetting.ExchangeAliasFieldValue.GetValue());
            }

            optionName = string.Empty;
            if (rbMailboxEnabled.Checked)
                optionName = "Mailbox-Enabled";
            else if (rbMailEnabled.Checked)
                optionName = "Mail-Enabled";

            if (optionName != AppSetting.InternalRecipientType.GetValue())
            {
                Guid appSettingsId = AppSetting.InternalRecipientType.Key;
                AppSetting.InternalRecipientType = new KeyValuePair<Guid, string>(appSettingsId, optionName);
                syncCore.crm.UpdateApplicationSetting(AppSetting.InternalRecipientType.Key, AppSetting.InternalRecipientType.GetValue());
            }

            optionName = string.Empty;
            if (rbExchangeAlias.Checked)
                optionName = "Alias";
            else if (rbEmailAddress.Checked)
                optionName = "EmailAddress";

            if (optionName != AppSetting.ExchangeIdentityType.GetValue())
            {
                Guid appSettingsId = AppSetting.ExchangeIdentityType.Key;
                AppSetting.ExchangeIdentityType = new KeyValuePair<Guid, string>(appSettingsId, optionName);
                syncCore.crm.UpdateApplicationSetting(AppSetting.ExchangeIdentityType.Key, AppSetting.ExchangeIdentityType.GetValue());
            }

            // Allow Mail User Update
            optionName = string.Empty;
            if (chkUpdateMailUser.Checked)
                optionName = "YES";
            else
                optionName = "NO";

            if (optionName != AppSetting.ExchangeMailUserUpdate.GetValue())
            {
                Guid appSettingsId = AppSetting.ExchangeMailUserUpdate.Key;
                AppSetting.ExchangeMailUserUpdate = new KeyValuePair<Guid, string>(appSettingsId, optionName);
                syncCore.crm.UpdateApplicationSetting(AppSetting.ExchangeMailUserUpdate.Key, AppSetting.ExchangeMailUserUpdate.GetValue());
            }

            // Logging/Trace Options
            StringBuilder value = new StringBuilder();

            if (chkInformation.Checked)
                value.Append("1");
            else
                value.Append("0");

            if (chkWarning.Checked)
                value.Append("1");
            else
                value.Append("0");

            if (chkError.Checked)
                value.Append("1");
            else
                value.Append("0");

            if (chkCritical.Checked)
                value.Append("1");
            else
                value.Append("0");

            if (value.ToString() != AppSetting.LoggingCriteria.GetValue())
            {
                Guid appSettingsId = AppSetting.LoggingCriteria.Key;
                AppSetting.LoggingCriteria = new KeyValuePair<Guid, string>(appSettingsId, value.ToString());
                syncCore.crm.UpdateApplicationSetting(AppSetting.LoggingCriteria.Key, AppSetting.LoggingCriteria.GetValue());
            }

            optionName = string.Empty;
            if (chkVerbose.Checked)
                optionName = "YES";
            else
                optionName = "NO";

            if (optionName != AppSetting.LoggingCriteriaVerbose.GetValue())
            {
                Guid appSettingsId = AppSetting.LoggingCriteriaVerbose.Key;
                AppSetting.LoggingCriteriaVerbose = new KeyValuePair<Guid, string>(appSettingsId, optionName);
                syncCore.crm.UpdateApplicationSetting(AppSetting.LoggingCriteriaVerbose.Key, AppSetting.LoggingCriteriaVerbose.GetValue());
            }

            optionName = string.Empty;
            if (chkRestartService.Checked)
                optionName = "YES";
            else
                optionName = "NO";

            if (optionName != AppSetting.RestartAutoSyncServicePostOperation.GetValue())
            {
                Guid appSettingsId = AppSetting.RestartAutoSyncServicePostOperation.Key;
                AppSetting.RestartAutoSyncServicePostOperation = new KeyValuePair<Guid, string>(appSettingsId, optionName);
                syncCore.crm.UpdateApplicationSetting(AppSetting.RestartAutoSyncServicePostOperation.Key, AppSetting.RestartAutoSyncServicePostOperation.GetValue());
            }

            optionName = string.Empty;
            if (chkDeleteOrphanedContacts.Checked)
                optionName = "YES";
            else
                optionName = "NO";

            if (optionName != AppSetting.DeleteExchangeContacts.GetValue())
            {
                Guid appSettingsId = AppSetting.DeleteExchangeContacts.Key;
                AppSetting.DeleteExchangeContacts = new KeyValuePair<Guid, string>(appSettingsId, optionName);
                syncCore.crm.UpdateApplicationSetting(AppSetting.DeleteExchangeContacts.Key, AppSetting.DeleteExchangeContacts.GetValue());
            }

            optionName = string.Empty;
            if (chkDeleteChangeTracking.Checked)
                optionName = "YES";
            else
                optionName = "NO";

            if (optionName != AppSetting.DeleteChangeTracking.GetValue())
            {
                Guid appSettingsId = AppSetting.DeleteChangeTracking.Key;
                AppSetting.DeleteChangeTracking = new KeyValuePair<Guid, string>(appSettingsId, optionName);
                syncCore.crm.UpdateApplicationSetting(AppSetting.DeleteChangeTracking.Key, AppSetting.DeleteChangeTracking.GetValue());
            }



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

        private void btnCancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void label39_Click(object sender, EventArgs e)
        {

        }

        private void label40_Click(object sender, EventArgs e)
        {

        }
    }
}
