using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Drawing;
using System.Data;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading;
using System.Windows.Forms;

namespace ExchangeSync
{
    public partial class AccountInfo : UserControl
    {
        public AccountInfo()
        {
            InitializeComponent();
        }

        private void AboutSettings_Load(object sender, EventArgs e)
        {
            ShowActivationInfo();
            // ShowProductInfo();
            // ShowDescription();
        }

        private void ShowActivationInfo()
        {
            if (!License.Evaluation)
            {
                // Licensed Version
                lblActivation.Text = "Product Licensed";
                lblRegisteredTo.Text = License.CompanyName;

                if (!string.IsNullOrEmpty(License.MaintenanceExpirationDate))
                {
                    DateTime maintenanceExpiresOn = Convert.ToDateTime(License.MaintenanceExpirationDate);
                    if (maintenanceExpiresOn >= DateTime.Today)
                        lblSupportExpiresOn.Text = "Support expires on " + License.MaintenanceExpirationDate;
                    else
                    {
                        lblSupportExpiresOn.Text = "Support agreement expired on " + License.MaintenanceExpirationDate;
                        lblSupportExpiresOn.ForeColor = Color.Red;
                    }
                }
                else
                {
                    lblSupportExpiresOn.Text = "Support agreement is not available";
                    lblSupportExpiresOn.ForeColor = Color.Red;
                }
            }
            else
            {
                lblActivation.Text = "Trial Version";
                lblActivation.ForeColor = Color.Red;

                if (License.EvaluationExpired)
                {
                    lblActivation.Text += " License Expired";
                }
                else
                {
                    lblActivation.Text += " License Will expire on " + License.ExpirationDate.ToShortDateString();
                }

                if (!string.IsNullOrEmpty(License.CompanyName))
                {
                    lblRegisteredTo.Text = "Trial Version Licensed to " + License.CompanyName;
                }

            }
        }

        private void btnAbout_Click(object sender, EventArgs e)
        {
            About frm = new About();
            frm.ShowDialog();
        }
    }
}
