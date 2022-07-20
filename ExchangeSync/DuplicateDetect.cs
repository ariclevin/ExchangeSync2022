using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

using Microsoft.Xrm.Sdk;

namespace ExchangeSync
{
    public partial class DuplicateDetect : Form
    {
        public DuplicateDetect()
        {
            InitializeComponent();
        }

        Core syncCore = ExchangeSyncManager.SyncCore;

        #region Duplicate Detection Functions

        private void DetectDuplicates()
        {
            DialogResult dr = MessageBox.Show("The duplicate detection process can be a length process, which uses the CRM duplicate detection rules on the contact entity. Are you sure you want to start the Duplicate Detection Process?", "Start Duplicate Detection", MessageBoxButtons.YesNo);
            if (dr == DialogResult.Yes)
            {
                bool useExistingList = false;
                if (AppSetting.DuplicateDetectionAction.Value == "EXISTING")
                    useExistingList = true;

                List<Guid> duplicates = syncCore.crm.CheckDuplicateEmails(useExistingList, 1);
                if (duplicates.Count > 0)
                {
                    string customfieldname = AppSetting.ContactAutoNumberFieldName.Value;
                    string revisionfieldname = AppSetting.ContactRevisionFieldName.Value;

                    // Clear Grid before adding fields
                    if (dgvDuplicates.Rows.Count > 0)
                    {
                        for (int i = dgvDuplicates.Rows.Count - 1; i >= 0; i--)
                        {
                            dgvDuplicates.Rows.RemoveAt(i);
                        }
                    }

                    foreach (Guid contactid in duplicates)
                    {
                        Entity contact = syncCore.crm.RetrieveContact(contactid);

                        int rowid = dgvDuplicates.Rows.Add();
                        DataGridViewRow row = (DataGridViewRow)dgvDuplicates.Rows[rowid];
                        row.Cells[0].Value = contactid.ToString();
                        row.Cells[1].Value = contact.Contains(customfieldname) ? contact.Attributes[customfieldname].ToString() : "";
                        row.Cells[2].Value = contact.Attributes["fullname"].ToString();
                        row.Cells[3].Value = contact.Contains("emailaddress1") ? contact.Attributes["emailaddress1"].ToString() : "";
                        row.Cells[4].Value = contact.Contains("xrm_exchangealias") ? contact.Attributes["xrm_exchangealias"].ToString() : "";
                        row.Cells[5].Value = true;
                        row.Cells[6].Value = contact.Contains(revisionfieldname) ? contact.Attributes[revisionfieldname].ToString() : "0";
                    }

                    if (!string.IsNullOrEmpty(customfieldname))
                    {
                        dgvDuplicates.Columns[1].Visible = true;
                        dgvDuplicates.Columns[1].Width = 100;
                    }
                }
                else
                    MessageBox.Show("No duplicate contact email addresses were found in CRM");
            }
        }

        private void UpdateExchangeAliases(bool clear)
        {
            for (int i = 0; i < dgvDuplicates.Rows.Count; i++)
            {
                DataGridViewRow row = dgvDuplicates.Rows[i];
                Guid id = new Guid(row.Cells[0].Value.ToString());
                string emailaddress = row.Cells[3].Value.ToString();
                int status = Convert.ToInt32(row.Cells[5].Value);
                if (status == 1)
                {
                    int revisionNo = Convert.ToInt32(row.Cells[6].Value);
                    if (!clear)
                    {
                        string alias = ExchangeHelper.GenerateNameFromEmailAddress(emailaddress);
                        syncCore.crm.UpdateContactAlias(id, emailaddress, revisionNo);
                        row.Cells[4].Value = alias;
                        row.Cells[4].ToolTipText = emailaddress;
                    }
                    else
                    {
                        syncCore.crm.UpdateContactAlias(id, "", revisionNo);
                        row.Cells[4].Value = "";
                    }
                }
            }
            // tslCurrent.Text = "Exchange Alias Update has been completed";

        }

        private void ChangeRowStatus()
        {
            if (dgvDuplicates.SelectedRows.Count > 0)
            {
                foreach (DataGridViewRow row in dgvDuplicates.SelectedRows)
                {
                    Guid rowId = new Guid(row.Cells[0].Value.ToString());
                    int status = Convert.ToInt32(row.Cells[5].Value);
                    if (status == 1)
                        row.Cells[5].Value = false;
                    else
                        row.Cells[5].Value = true;
                }
            }
            else
            {
                MessageBox.Show("Please select a row by clicking on the row selectors on the left", "No Rows Selected", MessageBoxButtons.OK);
            }

        }

        #endregion

    }
}
