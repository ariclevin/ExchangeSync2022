using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace ExchangeSync.Config
{
    public partial class BackupWizard : Form
    {
        public Core SyncCore { get; set; }
        public List<SyncEntity> exports = new List<SyncEntity>();
        public BackupWizard()
        {
            InitializeComponent();
        }

        private void BackupWizard_Load(object sender, EventArgs e)
        {
            List <KeyValuePair<string, string>> entities = SyncCore.crm.RetrieveAllEntities();
            cmbEntities.DataSource = entities;

        }

        private void btnBackup_Click(object sender, EventArgs e)
        {
            string key = ((DataRowView)cmbEntities.SelectedItem).Row[0].ToString();
            string value = ((DataRowView)cmbEntities.SelectedItem).Row[1].ToString();

            // RetrieveEntityValues
            // SyncCore.crm.RetrieveBusinessEntityCollection()
        }
    }
}
