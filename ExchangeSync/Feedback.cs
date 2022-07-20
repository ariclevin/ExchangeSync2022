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
    public partial class Feedback : UserControl
    {
        public Feedback()
        {
            InitializeComponent();
        }

        private void btnSubmit_Click(object sender, EventArgs e)
        {
            Messenger message = new Messenger();
            message.MailServer = ConfigSetting.ExchangeServerUrl;

            if (!string.IsNullOrEmpty(ConfigSetting.ExchangeCredentials.Domain))
            {
                message.UserName = ConfigSetting.ExchangeCredentials.Domain + "\\" + ConfigSetting.ExchangeCredentials.UserName;
            }
            else
            {
                message.UserName = ConfigSetting.ExchangeCredentials.UserName;
            }

            message.Password = ConfigSetting.ExchangeCredentials.Password;
            message.PortNumber = 25;

            string subject = String.Format("{0} [{1}]", txtSubject.Text, License.CompanyName);
            string from = "feedback@" + License.CompanyWebSite;
            string to = "exchangesyncsupport@briteglobal.com";

            message.CreateEmailMessage(from, to, string.Empty, string.Empty, subject, txtMessage.Text);
            int rc = message.SendEmail();
            message.Dispose();
        }
    }
}
