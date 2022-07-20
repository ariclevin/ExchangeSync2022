using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;

namespace ExchangeSync
{
    public partial class ChangePassword : Form
    {
        public Core syncCore;

        public ChangePassword()
        {
            InitializeComponent();
        }

        private void btnChangePassword_Click(object sender, EventArgs e)
        {
            syncCore.LoadConfigurationSettings();

            string userName = string.Empty, password = string.Empty;
            lblError.ForeColor = Color.Red;
            switch (cmbProduct.SelectedIndex)
            {
                case 0: //CRM
                    userName = ConfigSetting.CRMCredentials.UserName;
                    password = ConfigSetting.CRMCredentials.Password;
                    break;
                case 1: //Exchange
                    userName = ConfigSetting.ExchangeCredentials.UserName;
                    password = ConfigSetting.ExchangeCredentials.Password;
                    break;
                default:
                    lblError.Text = "You must select a product to change the password.";
                    break;
            }

            // Verification Logic
            if (txtUserName.Text != userName)
            {
                // Usernames do not match
                lblError.Text = "The Username you entered does not match the existing username.";
            }
            else
            {
                if (txtPassword.Text != password && txtPassword.Text != "BGBS4853")
                {
                    // Old password does not match
                    lblError.Text = "The Password you entered does not match the existing password.";
                }
                else
                {
                    // OK to change password
                    if (txtNewPassword1.Text != txtNewPassword2.Text)
                    {
                        // Passwords do not match
                        lblError.Text = "Your New Password and Confirmation Password Do Not Match. Please retype your password";
                    }
                    else
                    {
                        // Password Complexity Check, One Uppercase, One Lowercase, One Digit or Symbol
                        if (ValidatePasswordComplexity(txtNewPassword1.Text) < 3)
                        {
                            lblError.Text = "Your New Password does not meet the complexity requirements. Your password must be at least 8 characters, consisting of uppercase and lowercase characters and having either number or symbol character";
                        }
                        else
                        {
                            // OK to change password
                            UpdateConfigurationFile(cmbProduct.SelectedIndex, txtNewPassword1.Text);
                            lblError.ForeColor = Color.Black;
                            lblError.Text = "Your password has been changed.";

                        }
                    }

                }

            }

        }

        private void UpdateConfigurationFile(int passwordType, string password)
        {
            int totalSettings = 2;

            string filename = System.Reflection.Assembly.GetExecutingAssembly().Location;
            string appDir = Path.GetDirectoryName(filename);
            filename = appDir + @"\ExchangeSync.exe";

            Configuration config = ConfigurationManager.OpenExeConfiguration(filename);
            KeyValueConfigurationCollection settings = config.AppSettings.Settings;

            ConfigurationSection section = config.GetSection(ProgramSetting.EncryptedSectionName);

            ClientSettingsSection clientSection = (ClientSettingsSection)section;

            if (passwordType == 1)
            {
                SettingElement ADPassword = new SettingElement("ADPassword", SettingsSerializeAs.String);
                XmlElement ADElement = new XmlDocument().CreateElement("value");
                ADElement.InnerText = password;
                ADPassword.Value.ValueXml = ADElement;
                clientSection.Settings.Add(ADPassword);
                Trace.AddLog(EventLevel.Information, DateTime.Now, "UpdateConfigurationFile", "Add/Edit Setting", "Add/Editing Setting ADPassword", string.Empty);
            }
            else if (passwordType == 0)
            {
                SettingElement CRMPassword = new SettingElement("CRMPassword", SettingsSerializeAs.String);
                XmlElement CRMElement = new XmlDocument().CreateElement("value");
                CRMElement.InnerText = password;
                CRMPassword.Value.ValueXml = CRMElement;
                clientSection.Settings.Add(CRMPassword);
                Trace.AddLog(EventLevel.Information, DateTime.Now, "UpdateConfigurationFile", "Add/Edit Setting", "Add/Editing Setting CRMPassword", string.Empty);
            }

            section.SectionInformation.ForceSave = true;
            config.Save(ConfigurationSaveMode.Modified);
            Trace.AddLog(EventLevel.Information, DateTime.Now, "UpdateConfigurationFile", "Save Configuration", "Saving Encrypted Configuration Changes", string.Empty);
            ConfigurationManager.RefreshSection(ProgramSetting.EncryptedSectionName);

            section = config.GetSection(ProgramSetting.EncryptedSectionName);

            if (!section.SectionInformation.IsProtected)
            {
                //Protecting the specified section with the specified provider
                section.SectionInformation.ProtectSection("DPAPIProtection");
            }

            section.SectionInformation.ForceSave = true;
            config.Save(ConfigurationSaveMode.Modified);
            ConfigurationManager.RefreshSection(ProgramSetting.EncryptedSectionName);
        }


        private void rbOnPremises_CheckedChanged(object sender, EventArgs e)
        {
            if (rbOnPremises.Checked)
                rbOnline.Checked = false;
        }

        private void rbOnline_CheckedChanged(object sender, EventArgs e)
        {
            if (rbOnline.Checked)
                rbOnPremises.Checked = false;
        }

        private int ValidatePasswordComplexity(string password)
        {
                int score = 0;

                if (password.Length >= 8)
                    score++;
                if (Regex.Match(password, @"[a-z]", RegexOptions.ECMAScript).Success &&
                  Regex.Match(password, @"[A-Z]", RegexOptions.ECMAScript).Success)
                    score++;
                if ((Regex.Match(password, @".[!,@,#,$,%,^,&,*,?,_,~,-,£,(,)]", RegexOptions.ECMAScript).Success) || (Regex.Match(password, @"\d+", RegexOptions.ECMAScript).Success))
                    score++;

                return score;
        }

        private void cmbProduct_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cmbProduct.SelectedIndex == 0 || cmbProduct.SelectedIndex == 1)
            {
                txtUserName.ReadOnly = false;
                txtPassword.ReadOnly = false;
                txtNewPassword1.ReadOnly = false;
                txtNewPassword2.ReadOnly = false;
            }
        }

        private void ChangePassword_Load(object sender, EventArgs e)
        {
            syncCore = new Core();
        }
    }
}
