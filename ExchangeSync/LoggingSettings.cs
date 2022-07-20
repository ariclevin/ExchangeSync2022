using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace ExchangeSync
{
    public partial class LoggingSettings : UserControl
    {
        public LoggingSettings()
        {
            InitializeComponent();
        }

        private void LoggingSettings_Load(object sender, EventArgs e)
        {
            RetrieveApplicationSettings();
            DisableForm();            
        }

        private void RetrieveApplicationSettings()
        {
            if (!string.IsNullOrEmpty(AppSetting.LoggingCriteria.Value))
            {
                char[] criteria = AppSetting.LoggingCriteria.Value.ToCharArray();
                if (criteria.Length == 4)
                {
                    chkInformation.Checked = criteria[0] == '1' ? true : false;
                    chkWarning.Checked = criteria[1] == '1' ? true : false;
                    chkError.Checked = criteria[2] == '1' ? true : false;
                    chkCritical.Checked = criteria[3] == '1' ? true : false;
                }
            }

            if (!string.IsNullOrEmpty(AppSetting.LoggingCriteriaVerbose.Value))
            {
                string value = AppSetting.LoggingCriteriaVerbose.Value.ToLower();
                if (value == "yes")
                    chkVerbose.Checked = true;
                else 
                    chkVerbose.Checked = false;
            }
        }

        private void DisableForm()
        {
            chkVerbose.Enabled = false;
            chkInformation.Enabled = false;
            chkWarning.Enabled = false;
            chkError.Enabled = false;
            chkCritical.Enabled = false;
            btnSaveSettings.Enabled = false;
        }

        private void btnSaveSettings_Click(object sender, EventArgs e)
        {
            Core syncCore = ExchangeSyncManager.SyncCore;

            Guid key = AppSetting.LoggingCriteria.Key;
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

            AppSetting.LoggingCriteria = new KeyValuePair<Guid, string>(key, value.ToString());
            syncCore.crm.UpdateApplicationSetting(key, value.ToString());

            key = AppSetting.LoggingCriteriaVerbose.Key;
            if (chkVerbose.Checked)
            {
                AppSetting.LoggingCriteriaVerbose = new KeyValuePair<Guid, string>(key, "YES");
                syncCore.crm.UpdateApplicationSetting(key, "YES");
            }
            else
            {
                AppSetting.LoggingCriteriaVerbose = new KeyValuePair<Guid, string>(key, "NO");
                syncCore.crm.UpdateApplicationSetting(key, "NO");
            }
        }
    }
}
