using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Windows.Forms;

namespace ExchangeSync
{
    public partial class GeneralSettings : UserControl
    {
        public GeneralSettings()
        {
            InitializeComponent();
        }
        private bool PasswordChanged = false;
        private void ExchangeSettings_Load(object sender, EventArgs e)
        {
            RetrieveApplicationSettings();
        }

        private void RetrieveApplicationSettings()
        {
            switch (AppSetting.ListDeleteAction.Value.ToUpper())
            {
                case "REMOVECONNECTION":
                    chkRemoveConnection.Checked = true; chkRemoveConnection.Enabled = true;
                    chkDeleteExchangeList.Checked = false; chkDeleteExchangeList.Enabled = true;
                    chkDeleteOrphanedContacts.Checked = false; chkDeleteOrphanedContacts.Enabled = true;
                    break;
                case "DELETEGROUP":
                    chkRemoveConnection.Checked = true; chkRemoveConnection.Enabled = false;
                    chkDeleteExchangeList.Checked = true; chkDeleteExchangeList.Enabled = true;
                    chkDeleteOrphanedContacts.Checked = false; chkDeleteOrphanedContacts.Enabled = true;
                    break;
                case "DELETECONTACTS":
                    chkRemoveConnection.Checked = true; chkRemoveConnection.Enabled = false;
                    chkDeleteExchangeList.Checked = true; chkDeleteExchangeList.Enabled = false;
                    chkDeleteOrphanedContacts.Checked = true; chkDeleteOrphanedContacts.Enabled = true;
                    break;
            }

            if (AppSetting.DuplicateDetectionAction.Value.ToUpper() == "EXISTING")
                chkDuplicateDetectionAction.Checked = true;
            else
                chkDuplicateDetectionAction.Checked = false;

            if (AppSetting.DuplicateSyncAction.Value.ToUpper() == "MODIFY")
            {
                rbDuplicationSyncModify.Checked = true;
                rbDuplicationSyncIgnore.Checked = false;
            }
            else
            {
                rbDuplicationSyncModify.Checked = false;
                rbDuplicationSyncIgnore.Checked = true;
            }

            switch (AppSetting.ExchangeUpdateAction.Value.ToUpper())
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

            switch (AppSetting.HideFromExchangeAddressLists.Value.ToUpper())
            {
                case "HIDE":
                    rbHide.Checked = true;
                    rbShow.Checked = false;
                    rbIgnore.Checked = false;
                    break;
                case "SHOW":
                    rbHide.Checked = false;
                    rbShow.Checked = true;
                    rbIgnore.Checked = false;
                    break;
                case "IGNORE":
                    rbHide.Checked = false;
                    rbShow.Checked = false;
                    rbIgnore.Checked = true;
                    break;
                default:
                    rbHide.Checked = false;
                    rbShow.Checked = false;
                    rbIgnore.Checked = true;
                    break;
            }

            switch (AppSetting.FieldMappingAction.Value.ToUpper())
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

            switch (AppSetting.ExchangeMailContactFormat.Value.ToUpper())
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

            switch (AppSetting.ExchangeMailUserFormat.Value.ToUpper())
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
        }

        private void DisableForm()
        {
            foreach (Control ctrl in this.Controls)
            {
                ctrl.Enabled = false;
            }
        }

        private void btnSaveSettings_Click(object sender, EventArgs e)
        {
            Core syncCore = ExchangeSyncManager.SyncCore;

            // Save List Delete Action
            Guid key = AppSetting.ListDeleteAction.Key;
            if (chkDeleteOrphanedContacts.Checked)
            {
                AppSetting.ListDeleteAction = new KeyValuePair<Guid, string>(key, "DELETECONTACTS");
                syncCore.crm.UpdateApplicationSetting(key, "DELETECONTACTS");
            }
            else if (chkDeleteExchangeList.Checked)
            {
                AppSetting.ListDeleteAction = new KeyValuePair<Guid, string>(key, "DELETEGROUP");
                syncCore.crm.UpdateApplicationSetting(key, "DELETEGROUP");
            }
            else
            {
                AppSetting.ListDeleteAction = new KeyValuePair<Guid, string>(key, "REMOVECONNECTION");
                syncCore.crm.UpdateApplicationSetting(key, "REMOVECONNECTION");
            }

            // Save Duplicate Detection Action
            key = AppSetting.DuplicateDetectionAction.Key;
            if (chkDuplicateDetectionAction.Checked)
            {
                AppSetting.DuplicateDetectionAction = new KeyValuePair<Guid, string>(key, "EXISTING");
                syncCore.crm.UpdateApplicationSetting(key, "EXISTING");
            }
            else
            {
                AppSetting.DuplicateDetectionAction = new KeyValuePair<Guid, string>(key, "NEW");
                syncCore.crm.UpdateApplicationSetting(key, "NEW");
            }

            key = AppSetting.DuplicateSyncAction.Key;
            if (rbDuplicationSyncModify.Checked)
            {
                AppSetting.DuplicateSyncAction = new KeyValuePair<Guid, string>(key, "MODIFY");
                syncCore.crm.UpdateApplicationSetting(key, "MODIFY");
            }
            else if (rbDuplicationSyncIgnore.Checked)
            {
                AppSetting.DuplicateSyncAction = new KeyValuePair<Guid, string>(key, "IGNORE");
                syncCore.crm.UpdateApplicationSetting(key, "IGNORE");
            }

            // Save Exchange Update Action
            key = AppSetting.ExchangeUpdateAction.Key;
            if (rbUpdateAll.Checked)
            {
                AppSetting.ExchangeUpdateAction = new KeyValuePair<Guid, string>(key, "ALL");
                syncCore.crm.UpdateApplicationSetting(key, "ALL");
            }
            else if (rbUpdateDelta.Checked)
            {
                AppSetting.ExchangeUpdateAction = new KeyValuePair<Guid, string>(key, "DELTA");
                syncCore.crm.UpdateApplicationSetting(key, "DELTA");
            }

            // Save Default Hide Action
            key = AppSetting.HideFromExchangeAddressLists.Key;
            if (rbHide.Checked)
            {
                AppSetting.HideFromExchangeAddressLists = new KeyValuePair<Guid, string>(key, "HIDE");
                syncCore.crm.UpdateApplicationSetting(key, "HIDE");
            }
            else if (rbShow.Checked)
            {
                AppSetting.HideFromExchangeAddressLists = new KeyValuePair<Guid, string>(key, "SHOW");
                syncCore.crm.UpdateApplicationSetting(key, "SHOW");
            }
            else if (rbIgnore.Checked)
            {
                AppSetting.HideFromExchangeAddressLists = new KeyValuePair<Guid, string>(key, "IGNORE");
                syncCore.crm.UpdateApplicationSetting(key, "IGNORE");
            }

            // Save CRM Field Mapping Action
            key = AppSetting.FieldMappingAction.Key;
            if (rbMapNone.Checked)
            {
                AppSetting.FieldMappingAction = new KeyValuePair<Guid, string>(key, "NONE");
                syncCore.crm.UpdateApplicationSetting(key, "NONE");
            }
            else if (rbMapMinimal.Checked)
            {
                AppSetting.FieldMappingAction = new KeyValuePair<Guid, string>(key, "MIN");
                syncCore.crm.UpdateApplicationSetting(key, "MIN");
            }
            else if (rbMapCRM.Checked)
            {
                AppSetting.FieldMappingAction = new KeyValuePair<Guid, string>(key, "CRM");
                syncCore.crm.UpdateApplicationSetting(key, "CRM");
            }

            key = AppSetting.ExchangeMailContactFormat.Key;
            if (!string.IsNullOrEmpty(cmbMailContactFormat.Text))
            {
                string mailContactFormat = cmbMailContactFormat.Text.StripPunctuation().ToUpper();
                AppSetting.ExchangeMailContactFormat = new KeyValuePair<Guid,string>(key, mailContactFormat);
                syncCore.crm.UpdateApplicationSetting(key, mailContactFormat);
            }

            key = AppSetting.ExchangeMailUserFormat.Key;
            if (!string.IsNullOrEmpty(cmbMailUserFormat.Text))
            {
                string mailUserFormat = cmbMailUserFormat.Text.StripPunctuation().ToUpper();
                AppSetting.ExchangeMailContactFormat = new KeyValuePair<Guid, string>(key, mailUserFormat);
                syncCore.crm.UpdateApplicationSetting(key, mailUserFormat);
            }

            lblConfirmation.Text = "Application Settings have been saved.";
        }

        private void chkDeleteOrphanedContacts_CheckedChanged(object sender, EventArgs e)
        {
            // Check the other checkboxes as well and disable them
            chkDeleteExchangeList.Checked = chkDeleteOrphanedContacts.Checked;
            chkRemoveConnection.Checked = chkDeleteOrphanedContacts.Checked;
            chkDeleteExchangeList.Enabled = !chkDeleteOrphanedContacts.Checked;
            chkRemoveConnection.Enabled = chkDeleteOrphanedContacts.Checked;
        }

        private void chkDeleteExchangeList_CheckedChanged(object sender, EventArgs e)
        {
            chkRemoveConnection.Checked = chkDeleteOrphanedContacts.Checked;
            chkRemoveConnection.Enabled = chkDeleteOrphanedContacts.Checked;
        }

        private void cmbMailContactFormat_KeyDown(object sender, KeyEventArgs e)
        {
            
        }



    }
}
