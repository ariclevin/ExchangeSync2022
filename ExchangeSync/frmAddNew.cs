using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Xml;
using Microsoft.Xrm.Sdk;

namespace ExchangeSync
{
    public partial class frmAddNew : Form
    {
        private bool firstTime = true;

        public enum FormTypeCode
        {
            Create = 1,
            Update = 2
        }

        ExchangeSync.Core syncCore = ExchangeSyncManager.SyncCore;

        public FormTypeCode FormType { get; set; }

        public Guid SyncListId { get; set; }
        public Guid EntityId { get; set; } // Column 0
        public string EntityLogicalName { get; set; } // Column 1
        public string EntityDisplayName { get; set; } // Column 3
        public string EntityName { get; set; } // Column 2 (Group Name)
        public string ExchangeListName { get; set; } // Column 4 (Exchange Group Name)

        public frmAddNew()
        {
            InitializeComponent();
        }

        private void frmAddNew_Load(object sender, EventArgs e)
        {
            AddValuesToCRMGroupList();
            AddCRMListObjectTypes();
            AddExchangeDistributionLists();
            AddExchangeDistributionGroupTypes();

            if (FormType == FormTypeCode.Update)
            {
                lblTitle.Text = "Edit List";
                this.Text = "Edit";

                // Selection of Group Type               
                foreach (Infragistics.Win.UltraWinGrid.UltraGridRow row in ucGroupType.Rows)
                {
                    if (row.Cells[0].Value.ToString() == EntityLogicalName)
                    {
                        row.Selected = true;
                        ucGroupType.Enabled = false;
                        break;
                    }
                }


                foreach (object item in cmbCRMGroupName.Items)
                {
                    KeyValuePair<Guid, string> current = (KeyValuePair<Guid, string>)item;
                    if (current.Value == EntityDisplayName)
                    {
                        cmbCRMGroupName.SelectedItem = item;
                        cmbCRMGroupName.Text = EntityDisplayName;
                        cmbCRMGroupName.Enabled = false;
                    }
                }

                foreach (Infragistics.Win.UltraWinGrid.UltraGridRow row in ucDistributionList.Rows)
                {
                    if (row.Cells[0].Value.ToString() == ExchangeListName)
                    {
                        row.Selected = true;
                        break;
                    }
                }
            }
            else
            {
                lblTitle.Text = "Add New List";
                this.Text = "Add New";
            }
        }

        private void AddValuesToCRMGroupList()
        {
            List<CRMEntity> entities = new List<CRMEntity>();
            CRMEntity list = new CRMEntity();
            list.AddEntity("list", AppSetting.ListDisplayName.Value, "listname");
            entities.Add(list);

            EntityCollection intersects = syncCore.crm.RetrieveIntersects();
            if (intersects.Entities.Count > 0)
            {
                foreach (var result in intersects.Entities)
                {
                    string entityName = result["xrm_name"].ToString();
                    string displayName = result["xrm_displayname"].ToString();
                    string primaryAttributeName = result["xrm_primaryattributename"].ToString();

                    if (!string.IsNullOrEmpty(displayName))
                    {
                        CRMEntity crmEntity = new CRMEntity();
                        crmEntity.AddEntity(entityName, displayName, primaryAttributeName);
                        entities.Add(crmEntity);
                    }
                }
            }

            ucGroupType.DataSource = entities;
            ucGroupType.DisplayLayout.Bands[0].Columns[1].Width = ucGroupType.Width - 10;
            ucGroupType.DisplayLayout.Bands[0].Columns[0].Hidden = true;
            ucGroupType.DisplayLayout.Bands[0].Columns[2].Hidden = true;
        }

        private void AddExchangeOrganizationalUnits()
        {
            //List<OrganizationalUnit> lists = ExchangeHelper.utils.GetAllOrganizationalUnits();
            //ucDistributionList.DataSource = lists;
            //ucDistributionList.DisplayLayout.Bands[0].Columns[0].Width = (ucDistributionList.Width/2) - 10;
            //ucDistributionList.DisplayLayout.Bands[0].Columns[1].Width = (ucDistributionList.Width/2) + 10;
        }       
        
        private void AddExchangeDistributionLists()
        {
            List<DistributionGroup> lists = syncCore.exch.GetAllDistributionListGroups();
            ucDistributionList.DataSource = lists;
            ucDistributionList.DisplayLayout.Bands[0].Columns[0].Width = ucDistributionList.Width - 10;
            ucDistributionList.DisplayLayout.Bands[0].Columns[1].Hidden = true;
        }

        private void AddExchangeDistributionGroupTypes()
        {
            List<KeyValuePair<int, string>> list = new List<KeyValuePair<int, string>>();

            KeyValuePair<int, string> group1 = new KeyValuePair<int, string>(1, "Distribution");
            KeyValuePair<int, string> group2 = new KeyValuePair<int, string>(2, "Security");
            list.Add(group1); list.Add(group2);

            cmbExchangeGroupType.DataSource = list;
            cmbExchangeGroupType.ValueMember = "Key";
            cmbExchangeGroupType.DisplayMember = "Value";
        }

        private void AddCRMListObjectTypes()
        {
            List<KeyValuePair<int, string>> list = new List<KeyValuePair<int, string>>();

            KeyValuePair<int, string> group1 = new KeyValuePair<int, string>(1, "Account");
            KeyValuePair<int, string> group2 = new KeyValuePair<int, string>(2, "Contact");
            KeyValuePair<int, string> group3 = new KeyValuePair<int, string>(4, "Lead");
            list.Add(group1); list.Add(group2); list.Add(group3);

            cmbObjectType.DataSource = list;
            cmbObjectType.ValueMember = "Key";
            cmbObjectType.DisplayMember = "Value";

            cmbObjectType.SelectedIndex = 1;
            cmbObjectType.Enabled = false;
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btnOK_Click(object sender, EventArgs e)
        {
            bool rc = true;
            
            Infragistics.Win.UltraWinGrid.UltraGridRow row = ucGroupType.SelectedRow;
            string entityname = "", entitydisplayname = "", attributename = "";
            if (row != null)
            {
                if (!row.IsEmptyRow)
                {
                    entityname = row.Cells[0].Value.ToString();
                    entitydisplayname = row.Cells[1].Value.ToString();
                    attributename = row.Cells[2].Value.ToString();
                }
                else
                    rc = false;
            }
            else
                rc = false;

            string listId = cmbCRMGroupName.SelectedValue.ToString();
            string listName = cmbCRMGroupName.Text.ToString();

            string exchangelistname = "";
            if (ucDistributionList.Enabled)
            {
                if (ucDistributionList.SelectedRow != null)
                {
                    if (!ucDistributionList.SelectedRow.IsEmptyRow)
                        exchangelistname = ucDistributionList.Text.ToString();
                    else
                        rc = false;
                }
                else
                    rc = false;
            }
            else
            {
                if (txtOther.Text != "")
                {
                    exchangelistname = txtOther.Text;
                    string ou = AppSetting.ExchangeDistributionGroupOU.Value;
                    if (cmbExchangeGroupType.SelectedItem != null)
                    {
                        KeyValuePair<int, string> selection = (KeyValuePair<int, string>)cmbExchangeGroupType.SelectedItem;
                        syncCore.exch.CreateDistributionGroup(exchangelistname, ou, (GroupType)selection.Key);
                    }
                    else
                    {
                        syncCore.exch.CreateDistributionGroup(exchangelistname, ou);
                    }
                    
                }
                else
                    rc = false;
            }


            if (rc)
            {
                if (FormType == FormTypeCode.Create)
                {
                    if (cmbExchangeGroupType.SelectedItem != null)
                    {
                        KeyValuePair<int, string> selection = (KeyValuePair<int, string>)cmbExchangeGroupType.SelectedItem;
                        KeyValuePair<int, string> objectType = (KeyValuePair<int, string>)cmbObjectType.SelectedItem;
                        syncCore.crm.CreateSyncList(ProgramSetting.ApplicationProfile.Value, entityname, entitydisplayname, new Guid(listId), listName, exchangelistname, selection.Key, objectType.Key);
                    }
                    else
                    {
                        syncCore.crm.CreateSyncList(ProgramSetting.ApplicationProfile.Value, entityname, entitydisplayname, new Guid(listId), listName, exchangelistname, 1, 2);
                    }

                }
                else if (FormType == FormTypeCode.Update)
                {
                    syncCore.crm.UpdateSyncList(SyncListId, exchangelistname);
                }
                this.Close();
            }
            else
            {
                MessageBox.Show("Please make sure that you have selected the CRM Group Type and Group Name and the Exchange List Name (or New List Name).\nCorrect your errors are try again");
            }
        }

        private void btnNew_Click(object sender, EventArgs e)
        {
            txtOther.Enabled = true;
            txtOther.ForeColor = System.Drawing.Color.Black;
            ucDistributionList.Enabled = false;
            btnNew.Enabled = false;
            txtOther.Focus();
            txtOther.SelectAll();
        }

        private void txtOther_TextChanged(object sender, EventArgs e)
        {

        }

        private void ucGroupType_RowSelected2(object sender, Infragistics.Win.UltraWinGrid.RowSelectedEventArgs e)
        {
            Core syncCore = ExchangeSyncManager.SyncCore;

            if (e.Row.Cells[0].Value.GetType() == System.Type.GetType("System.String"))
            {
                // int selectedValue = Convert.ToInt32(cmbGroupType.SelectedValue);
                string entityname = e.Row.Cells[0].Value.ToString();
                string attributename = e.Row.Cells[2].Value.ToString();

                // if (selectedValue != 0)
                EntityCollection results = new EntityCollection();

                try
                {
                    results = syncCore.crm.RetrieveSyncLists(ProgramSetting.ApplicationProfile.Value, entityname);
                }
                catch (System.Exception ex)
                {
                    MessageBox.Show("The following error occured trying to retrieve the lists:\n" + ex.Message);
                }

                if (entityname != "")
                {
                    List<KeyValuePair<Guid, string>> list = new List<KeyValuePair<Guid, string>>();
                    foreach (var result in results.Entities)
                    {
                        KeyValuePair<Guid, string> reference = new KeyValuePair<Guid, string>(result.Id, result.Attributes[attributename].ToString());
                        list.Add(reference);
                    }

                    cmbCRMGroupName.DataSource = list;
                    cmbCRMGroupName.ValueMember = "Key";
                    cmbCRMGroupName.DisplayMember = "Value";
                }
            }
            else
                firstTime = false;

        }

        
        private void ucGroupType_RowSelected(object sender, Infragistics.Win.UltraWinGrid.RowSelectedEventArgs e)
        {
            Core syncCore = ExchangeSyncManager.SyncCore;

            if (e.Row.Cells[0].Value.GetType() == System.Type.GetType("System.String"))
            {
                // int selectedValue = Convert.ToInt32(cmbGroupType.SelectedValue);
                string entityname = e.Row.Cells[0].Value.ToString();
                string attributename = e.Row.Cells[2].Value.ToString();

                // if (selectedValue != 0)
                EntityCollection results = new EntityCollection();
                string textfield = "";
                if (entityname != "")
                {
                    switch (entityname)
                    {
                        case "list":
                            try
                            {
                                results = syncCore.crm.RetrieveLists(false);
                            }
                            catch (System.Exception ex)
                            {
                                MessageBox.Show("The following error occured trying to retrieve the lists:\n" + ex.Message);
                            }
                            textfield = "listname";
                            break;
                        default:
                            // Need to get primaryAttributeName
                            results = syncCore.crm.RetrieveBusinessEntityCollection(entityname, attributename, false);
                            textfield = attributename;
                            break;
                    }
                }

                if (entityname != "")
                {
                    List<KeyValuePair<Guid, string>> list = new List<KeyValuePair<Guid, string>>();
                    foreach (var result in results.Entities)
                    {
                        KeyValuePair<Guid, string> reference = new KeyValuePair<Guid, string>(result.Id, result.Attributes[textfield].ToString());
                        list.Add(reference);
                    }

                    cmbCRMGroupName.DataSource = list;
                    cmbCRMGroupName.ValueMember = "Key";
                    cmbCRMGroupName.DisplayMember = "Value";
                }
            }
            else
                firstTime = false;

        }

        private void ucGroupType_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
        {

        }
    }

}
