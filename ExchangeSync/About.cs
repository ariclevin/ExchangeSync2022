using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Windows.Forms;

namespace ExchangeSync
{
    public partial class About : Form
    {
        public About()
        {
            InitializeComponent();
        }

        private void ShowProductInfo()
        {
            string productName = String.Format("CRM Exchange Sync {0} ({1}-bit)", Assembly.GetExecutingAssembly().GetName().Version, "32");
            lblProductName.Text = productName;
        }

        private void ShowDescription()
        {
            StringBuilder sb = new StringBuilder();
            sb.AppendLine("This product contains components and third-party utilities from Microsoft Corporation and Infragistics, Inc.");
            sb.AppendLine("");
            sb.AppendLine("Microsoft Dynamics CRM, Microsoft Exchange and Microsoft Visual Studio are trademarks of Microsoft Corporation.");
            sb.AppendLine("");
            sb.AppendLine("Infragistics and NetAdvantage are trademarks of Infragistics, Inc.");
            txtDescription.Text = sb.ToString();
        }

        private void About_Load(object sender, EventArgs e)
        {
            ShowProductInfo();
            ShowDescription();
        }
    }
}
