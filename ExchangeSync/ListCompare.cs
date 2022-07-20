using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using ExchangeSync.Properties;
using Infragistics.Win.UltraWinGrid;
using Infragistics.Win.UltraWinToolbars;
using Microsoft.Xrm.Sdk;


namespace ExchangeSync
{
    public partial class ListCompare : Form
    {
        ExchangeSync.Core syncCore = ExchangeSyncManager.SyncCore;

        public Guid CRMListId { get; set; }
        public string CRMEntityName { get; set; }
        public string CRMListName { get; set; }
        public string ExchangeListName { get; set; }

        public bool CompareAll { get; set; }

        public ListCompare()
        {
            CompareAll = true;
            InitializeComponent();
        }

        public ListCompare(Guid crmListId, string crmEntityName, string crmListName, string exchangeListName)
        {
            CRMListId = crmListId;
            CRMEntityName = crmEntityName;
            CRMListName = crmListName;
            ExchangeListName = exchangeListName;
            CompareAll = false;

            InitializeComponent();
        }

        private void ListCompare_Load(object sender, EventArgs e)
        {
            InitializeData();
        }

        private void InitializeData()
        {
            List<Recipient> list = new List<Recipient>();

            if (!CompareAll)
            {
                EntityCollection contacts = syncCore.RetrieveContacts(CRMEntityName, CRMListName, CRMListId);
                list = CompareSingleList(contacts, list);
            }
            else
            {
                EntityCollection syncLists = syncCore.crm.RetrieveSyncLists(ProgramSetting.ApplicationProfile.Value, "");
                if (syncLists.Entities.Count > 0)
                {
                    foreach (Entity syncList in syncLists.Entities)
                    {
                        Guid syncListId = syncList.Id;
                        Guid listId = new Guid(syncList.GetAttributeValue<string>("xrm_entityid"));
                        string listName = syncList.GetAttributeValue<string>("xrm_listname");
                        string entityName = syncList.GetAttributeValue<string>("xrm_entityname");
                        bool isActive = syncList.GetAttributeValue<bool>("xrm_liststatus");

                        if (isActive)
                        {
                            EntityCollection contacts = syncCore.RetrieveContacts(entityName, listName, listId);
                            if (contacts.Entities.Count > 0)
                            {
                                foreach (Entity contact in contacts.Entities)
                                {
                                    Guid contactId = contact.Id;
                                    string emailAddress = contact.GetAttributeValue<string>("emailaddress1");
                                    string alias = contact.GetAttributeValue<string>("xrm_exchangealias");
                                    string fullName = contact.GetAttributeValue<string>("fullname");
                                    int stateCode = contact.GetAttributeValue<OptionSetValue>("statecode").Value;
                                    string status = stateCode > 0 ? "Inactive" : "Active";

                                    if (!list.Exists(x => x.EmailAddress == emailAddress.ToLower()))
                                    {
                                        if (status != "Inactive" && AppSetting.CompareInactive.ToString().ToLower() != "no")
                                            list.Add(new Recipient(contactId, emailAddress, alias, fullName, status));
                                    }
                                }
                            }
                        }
                    }
                }

                // Run Comparison Against Exchange
                foreach (Recipient r in list)
                {
                    string emailAddress = r.EmailAddress;
                    emailAddress = emailAddress.IsValidEmail() ? emailAddress.ToLower() : "";

                    string crmAlias = r.CrmAlias;

                    // Check is this is Mail Contact or Mail User
                    try
                    {
                        ExchangeContact ec = syncCore.exch.GetRecipient(emailAddress);
                        if (ec.Alias.Equals(crmAlias, StringComparison.OrdinalIgnoreCase))
                        {
                            // Matching Alias
                            r.ExchangeAlias = ec.Alias;
                            r.Message = "CRM and Exchange Recipients Match";
                            r.Success = true;
                        }
                        else
                        {
                            // Alias does not match
                            r.ExchangeAlias = ec.Alias;
                            r.Message = "CRM and Exchange Recipients Do Not Match";
                            r.Success = false;
                        }
                    }
                    catch (System.Exception ex)
                    {
                        // Alias does not match
                        r.Message = ex.Message;
                        r.Success = false;
                    }
                }
            }
            ugCompareList.DataSource = list;
        }

        private List<Recipient> CompareSingleList(EntityCollection contacts, List<Recipient> list)
        {
            if (contacts.Entities.Count > 0)
            {
                foreach (Entity contact in contacts.Entities)
                {
                    Guid contactId = contact.Id;
                    string emailAddress = contact.GetAttributeValue<string>("emailaddress1");
                    string alias = contact.GetAttributeValue<string>("xrm_exchangealias");
                    string fullName = contact.GetAttributeValue<string>("fullname");
                    int stateCode = contact.GetAttributeValue<OptionSetValue>("statecode").Value;
                    string status = stateCode > 0 ? "Inactive" : "Active";
                    if (status != "Inactive" && AppSetting.CompareInactive.ToString().ToLower() != "no")
                    {
                        list.Add(new Recipient(contactId, emailAddress, alias, fullName, status));
                    }
                        
                }
            }

            List<ExchangeContact> recipients = syncCore.exch.GetDistributionListMembers(ExchangeListName);

            if (recipients.Count > 0)
            {
                foreach (ExchangeContact contact in recipients)
                {
                    string emailAddress = contact.PrimarySmtpAddress;
                    emailAddress = emailAddress.IsValidEmail() ? emailAddress.ToLower() : "";

                    string alias = contact.Alias;

                    if (list.Exists(x => x.EmailAddress == emailAddress))
                    {
                        Recipient recipient = list.Find(x => x.EmailAddress == emailAddress);
                        recipient.ExchangeAlias = alias;
                        recipient.Message = "CRM and Exchange Recipients Match";
                        recipient.Success = true;
                    }

                }
            }

            foreach (Recipient r in list)
            {
                if (r.Message == "Comparison Pending")
                {
                    r.Message = "Could Not Find Exchange Recipient";
                    r.Success = false;
                }
            }

            return list;
        }



        private void ugCompareList_InitializeRow(object sender, Infragistics.Win.UltraWinGrid.InitializeRowEventArgs e)
        {
            Color color = Color.Black;
            switch (Convert.ToBoolean(e.Row.Cells["Success"].Value))
            {
                case true:
                    break;
                case false:
                    color = Color.Red;
                    break;
                default:
                    break;
            }

            for (int i = 0; i < e.Row.Cells.Count; i++)
            {
                e.Row.Cells[i].Appearance.ForeColor = color;
            }
        }

        private void ugCompareList_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
        {
            e.Layout.Bands[0].Columns["RecipientId"].Hidden = true;
            e.Layout.Bands[0].Columns["FullName"].Width = 150;
            e.Layout.Bands[0].Columns["EmailAddress"].Width = 200;
            e.Layout.Bands[0].Columns["CrmAlias"].Width = 150;
            e.Layout.Bands[0].Columns["ExchangeAlias"].Width = 150;
            e.Layout.Bands[0].Columns["Message"].Width = 300;
            e.Layout.Bands[0].Columns["Status"].Width = 100;
            e.Layout.Bands[0].Columns["Success"].Hidden = true;
        }

        private void utmListCompare_ToolClick(object sender, Infragistics.Win.UltraWinToolbars.ToolClickEventArgs e)
        {
            switch (e.Tool.Key)
            {
                case "toolShowAll":
                    CheckTool(e.Tool.Key);
                    ShowRows(true, true);
                    break;

                case "toolShowActive":
                    CheckTool(e.Tool.Key);
                    ShowRows(true, false);
                    break;

                case "toolShowInactive":
                    CheckTool(e.Tool.Key);
                    ShowRows(false, true);
                    break;

                case "toolRefresh":
                    InitializeData();
                    break;
            }
        }

        private void CheckTool(string key)
        {
            StateButtonTool showAll = (StateButtonTool)utmListCompare.Tools["toolShowAll"];
            showAll.Checked = showAll.Key == key;

            StateButtonTool showActive = (StateButtonTool)utmListCompare.Tools["toolShowActive"];
            showActive.Checked = showActive.Key == key;

            StateButtonTool showInactive = (StateButtonTool)utmListCompare.Tools["toolShowInactive"];
            showInactive.Checked = showInactive.Key == key;
        }

        private void ShowRows(bool showActive, bool showInactive)
        {
            foreach (UltraGridRow row in ugCompareList.Rows.All)
            {
                string rowStatus = row.Cells["Status"].Text;
                switch (rowStatus)
                {
                    case "Active":
                        row.Hidden = !showActive;
                        break;
                    case "Inactive":
                        row.Hidden = !showInactive;
                        break;
                }
            }
        }

        private void btnResetExchangeAlias_Click(object sender, EventArgs e)
        {
            foreach (UltraGridRow row in ugCompareList.Selected.Rows)
            {
                string emailAddress = row.Cells["EmailAddress"].Text;
                string crmAlias = row.Cells["CrmAlias"].Text;
                string exchangeAlias = row.Cells["ExchangeAlias"].Text;

                syncCore.exch.UpdateMailContactAlias(emailAddress, crmAlias);
                row.Cells["ExchangeAlias"].SetValue(crmAlias, false);
            }
        }

        private void btnResetCrmAlias_Click(object sender, EventArgs e)
        {
            foreach (UltraGridRow row in ugCompareList.Selected.Rows)
            {
                
                string contactId = row.Cells["RecipientId"].Text;
                string emailAddress = row.Cells["EmailAddress"].Text;
                string crmAlias = row.Cells["CrmAlias"].Text;
                string exchangeAlias = row.Cells["ExchangeAlias"].Text;

                syncCore.crm.UpdateContactAlias(new Guid(contactId), exchangeAlias);
                row.Cells["CrmAlias"].SetValue(exchangeAlias, false);
            }
        }
    }


    public class Recipient
    {
        public Guid RecipientId { get; set; }
        public string EmailAddress { get; set; }
        public string CrmAlias { get; set; }
        public string FullName { get; set; }
        public string ExchangeAlias { get; set; }
        public string Message { get; set; }
        public string Status { get; set; }
        public bool Success { get; set; }

        public Recipient (Guid recipientId, string emailAddress, string fullName, string status)
        {
            RecipientId = recipientId;
            EmailAddress = emailAddress.IsValidEmail() ? emailAddress.ToLower() : "";
            FullName = fullName;
            Message = "Comparison Pending";
            Status = status;
            Success = false;
        }
        public Recipient(Guid recipientId, string emailAddress, string crmAlias, string fullName, string status)
        {
            RecipientId = recipientId;
            EmailAddress = emailAddress.IsValidEmail() ? emailAddress.ToLower() : "";
            CrmAlias = crmAlias;
            FullName = fullName;
            Message = "Comparison Pending";
            Status = status;
            Success = false;
        }

    }
}
