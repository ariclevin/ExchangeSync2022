using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

using Infragistics.Win.Misc;

using Microsoft.Xrm.Sdk;

namespace ExchangeSync
{
    public partial class ListManager : Form
    {

        Core syncCore = ExchangeSyncManager.SyncCore;

        public ListManager()
        {
            InitializeComponent();
        }

        private void ListManager_Load(object sender, EventArgs e)
        {

            // Required Attributes Field
            // PopulateDisconnectedLists();
        }

        private void ImportExistingLists()
        {
            if (dgvDisconnected.Rows.Count > 0)
            {
                for (int i = dgvDisconnected.Rows.Count - 1; i >= 0; i--)
                {
                    dgvDisconnected.Rows.RemoveAt(i);
                }
            }

            List<KeyValuePair<Guid, string>> syncList = RetrieveSyncLists();

            EntityCollection results = syncCore.crm.RetrieveLists(true); // Lists that have an exchangedlname
            if (results.Entities.Count > 0)
            {
                foreach (var result in results.Entities)
                {
                    Guid listId = result.Id;
                    if (!syncList.Contains(listId))
                    {
                        int rowid = dgvDisconnected.Rows.Add();
                        DataGridViewRow row = (DataGridViewRow)dgvDisconnected.Rows[rowid];
                        row.Cells[0].Value = listId.ToString();
                        row.Cells[1].Value = "list";
                        row.Cells[2].Value = AppSetting.ListDisplayName.Value;
                        row.Cells[3].Value = result.Attributes["listname"].ToString();
                        row.Cells[4].Value = result.Contains("xrm_exchangedlname") ? result.Attributes["xrm_exchangedlname"].ToString() : string.Empty;
                        OptionSetValue stateCode = ((OptionSetValue)(result.Attributes["statecode"]));
                        if (stateCode.Value == 0)
                            row.Cells[5].Value = true;
                    
                    }                   
                }
            }

        }

        private List<KeyValuePair<Guid, string>> RetrieveSyncLists()
        {
            List<KeyValuePair<Guid, string>> syncList = new List<KeyValuePair<Guid, string>>();
            EntityCollection results = syncCore.crm.RetrieveSyncLists(ProgramSetting.ApplicationProfile.Value);
            if (results.Entities.Count > 0)
            {
                foreach (Entity result in results.Entities)
                {
                    Guid listId = new Guid(result["xrm_entityid"].ToString());
                    string listName = result["xrm_listname"].ToString();
                    syncList.Add(new KeyValuePair<Guid, string>(listId, listName));
                }
            }

            return syncList;
        }

        private void PopulateDisconnectedLists()
        {
            if (dgvDisconnected.Rows.Count > 0)
            {
                for (int i = dgvDisconnected.Rows.Count - 1; i >= 0; i--)
                {
                    dgvDisconnected.Rows.RemoveAt(i);
                }
            }

            List<KeyValuePair<Guid, string>> syncList = RetrieveSyncLists();

            EntityCollection results = syncCore.crm.RetrieveDisconnectedLists();
            if (results.Entities.Count > 0)
            {
                foreach (var result in results.Entities)
                {
                    Guid listId = result.Id;
                    if (!syncList.Contains(listId))
                    {
                        int rowid = dgvDisconnected.Rows.Add();
                        DataGridViewRow row = (DataGridViewRow)dgvDisconnected.Rows[rowid];
                        row.Cells[0].Value = listId.ToString();
                        row.Cells[1].Value = "list";
                        row.Cells[2].Value = AppSetting.ListDisplayName.Value;
                        row.Cells[3].Value = result.Attributes["listname"].ToString();
                        OptionSetValue stateCode = ((OptionSetValue)(result.Attributes["statecode"]));
                        if (stateCode.Value == 0)
                            row.Cells[5].Value = true;
                    }
                }
            }
        }

        private void PopulateExchangeGroupNames()
        {
            if (dgvDisconnected.Rows.Count > 0)
            {
                foreach (DataGridViewRow row in dgvDisconnected.Rows)
                {
                    if (row.Selected)
                    {
                        string listName = row.Cells[3].Value.ToString();
                        string exchangeListName = ExchangeHelper.TrimInvalidCharacters(listName);
                        exchangeListName = exchangeListName.Replace(" ", "_");
                        row.Cells[4].Value = exchangeListName;
                    }
                }
            }
        }

        private void utmList_ToolClick(object sender, Infragistics.Win.UltraWinToolbars.ToolClickEventArgs e)
        {
            switch (e.Tool.Key)
            {
                case "toolPopulateExchangeGroupNames":
                    PopulateExchangeGroupNames();
                    break;
                case "toolImportNewLists":
                    PopulateDisconnectedLists();
                    break;
                case "toolImportExistingLists":
                    ImportExistingLists();
                    break;
            }
        }

        private void btnOK_Click(object sender, EventArgs e)
        {
            if (dgvDisconnected.Rows.Count > 0)
            {
                foreach (DataGridViewRow row in dgvDisconnected.SelectedRows)
                {
                    Guid crmListId = new Guid(row.Cells[0].Value.ToString());
                    string exchangeGroupName = row.Cells[4].Value.ToString();

                    bool success = false;
                    bool groupExists = false;
                    try
                    {
                        groupExists = syncCore.exch.DistributionGroupExists(exchangeGroupName);
                    }
                    catch (System.Exception ex)
                    {
                        success = false;
                        string message = ex.Message;
                        string command = ex.InnerException.Data["CommandName"].ToString();
                        string function = ex.InnerException.Data["MethodName"].ToString();
                        List<KeyValuePair<string, string>> parameters = (List<KeyValuePair<string, string>>)ex.InnerException.Data["Parameters"];
                        string sParams = Helper.KeyValuePairListToString(parameters);

                        Trace.AddLog(EventLevel.Warning, DateTime.Now, "ListManager", function, command, message, "Validation of Distribution Group in Exchange via List Manager", sParams);
                    }

                    if (!groupExists)
                    {
                        try
                        {
                            success = syncCore.exch.CreateDistributionGroup(exchangeGroupName, AppSetting.ExchangeDistributionGroupOU.Value);
                        }
                        catch (System.Exception ex)
                        {
                            success = false;
                            string message = ex.Message;
                            string command = ex.InnerException.Data["CommandName"].ToString();
                            string function = ex.InnerException.Data["MethodName"].ToString();
                            List<KeyValuePair<string, string>> parameters = (List<KeyValuePair<string, string>>)ex.InnerException.Data["Parameters"];
                            string sParams = Helper.KeyValuePairListToString(parameters);

                            Trace.AddLog(EventLevel.Warning, DateTime.Now, "ListManager", function, command, message, "Addition of Distribution Group to Exchange via List Manager", sParams);
                        }
                    }
                    else
                        success = true;

                    if (success)
                    {
                        Guid listId = new Guid(row.Cells[0].Value.ToString());
                        string entityName = row.Cells[1].Value.ToString();
                        string entityDisplayName = row.Cells[2].Value.ToString();
                        string listName = row.Cells[3].Value.ToString();
                        
                        Guid syncListId = syncCore.crm.CreateSyncList(ProgramSetting.ApplicationProfile.Value, entityName, entityDisplayName, listId, listName, exchangeGroupName);

                        UltraDesktopAlert alert = new UltraDesktopAlert(this.components);
                        if (syncListId != Guid.Empty)
                        {
                            UltraDesktopAlertShowWindowInfo info = new UltraDesktopAlertShowWindowInfo("Completed", String.Format("The list {0} has been added to Exchange and linked to CRM.", listName));
                            alert.Show(info);
                        }
                        else
                        {
                            UltraDesktopAlertShowWindowInfo info = new UltraDesktopAlertShowWindowInfo("Error", String.Format("The was an error processing the list {0}.", listName));
                            alert.Show(info);
                        }

                    }
                }
            }


            this.Close();
            
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {

        }
    }
}
