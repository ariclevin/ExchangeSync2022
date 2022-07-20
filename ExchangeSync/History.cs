using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace ExchangeSync
{
    public partial class History : UserControl
    {
        public History()
        {
            InitializeComponent();
        }

        private void Button_MouseHover(object sender, EventArgs e)
        {
            Button btn = sender as Button;
            btn.BackColor = Color.LightSkyBlue;

        }

        private void Button_MouseLeave(object sender, EventArgs e)
        {
            Button btn = sender as Button;
            btn.BackColor = Color.Transparent;
        }

        private void btnRecent_Click(object sender, EventArgs e)
        {
            DirectoryInfo di = new DirectoryInfo(@"C:\logs");
            FileInfo[] files = di.GetFiles("*.csv");

            foreach (FileInfo file in files)
            {
                // NEED TO TEST
                ulvFiles.Items.Add(new Infragistics.Win.UltraWinListView.UltraListViewItem(file.Name));

            }

        }

        private void btnExchangeSync_Click(object sender, EventArgs e)
        {

        }
    }
}
