using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading;
using System.Windows.Forms;

namespace ExchangeSync
{
    public partial class Password : Form
    {
        public string AppConfigFileName { get; set; }
        string locale = "";

        public Password()
        {
            InitializeComponent();

        }

        private void btnSubmitChanges_Click(object sender, EventArgs e)
        {
            if (txtNewPassword.Text == txtConfirmPassword.Text)
            {
                if (IsValidLicense())
                {
                    UpdatePassword();
                    MessageBox.Show("The change password has been completed. Please restart the application for the changes to be refreshed", "Change Password", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    this.Close();
                }
                else
                    MessageBox.Show("The change password feature is only available for production licenses", "Non-Licensed version", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                MessageBox.Show("Your new passwords do not match, please re-enter and try again", "Password do not match", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
            }
        }

        private void UpdatePassword()
        {
            string filename = System.Reflection.Assembly.GetExecutingAssembly().Location;
            string directory = System.IO.Path.GetDirectoryName(filename);
            AppConfigFileName = directory + @"\CRMExchangeSync.exe";

            string machineName = Network.MachineName;
            string ipAddress = Network.IPAddress;

            UpdateApplicationSettings(ProgramSetting.EncryptedSectionName, txtNewPassword.Text);
        }

        private void UpdateApplicationSettings(string sectionName, string password, string liveIdPassword = "")
        {
            Configuration config = ConfigurationManager.OpenExeConfiguration(AppConfigFileName);
            ConfigurationSection section = config.GetSection(sectionName);

            ClientSettingsSection clientSection = (ClientSettingsSection)section;

            SettingElement setPassword = clientSection.Settings.Get("ADPassword");
            setPassword.Value.ValueXml.InnerXml = password;

            if (!string.IsNullOrEmpty(liveIdPassword))
            {
                SettingElement setLiveIdPassword = clientSection.Settings.Get("CRMPassword");
                setLiveIdPassword.Value.ValueXml.InnerXml = liveIdPassword;
            }

            section.SectionInformation.ForceSave = true;
            config.Save(ConfigurationSaveMode.Modified);
            ConfigurationManager.RefreshSection(sectionName);
        }

        private void ProtectSection(string sectionName, string provName)
        {
            Configuration config = ConfigurationManager.OpenExeConfiguration(AppConfigFileName);
            ConfigurationSection section = config.GetSection(sectionName);

            if (!section.SectionInformation.IsProtected)
            {
                //Protecting the specified section with the specified provider
                section.SectionInformation.ProtectSection(provName);
            }

            section.SectionInformation.ForceSave = true;
            config.Save(ConfigurationSaveMode.Modified);
            ConfigurationManager.RefreshSection(sectionName);
        }

        private string GetWindowsLiveIdPassword(string sectionName, string provName)
        {
            Configuration config = ConfigurationManager.OpenExeConfiguration(AppConfigFileName);
            ConfigurationSection section = config.GetSection(sectionName);

            if (!section.SectionInformation.IsProtected)
            {
                //Protecting the specified section with the specified provider
                section.SectionInformation.UnprotectSection();
            }

            section.SectionInformation.ForceSave = true;
            config.Save(ConfigurationSaveMode.Modified);
            ConfigurationManager.RefreshSection(sectionName);

            ClientSettingsSection clientSection = (ClientSettingsSection)section;

            SettingElement setLiveIdPassword = clientSection.Settings.Get("LiveIdPassword");
            string liveIdPassword = setLiveIdPassword.Value.ValueXml.InnerXml.ToString();
            return liveIdPassword;
        }



        private bool IsValidLicense()
        {
            bool rc = false;
            string embeddedLicense = "";
            if (!License.LicenseKey.Contains("watermark"))
                embeddedLicense = License.LicenseKey;

            if (!string.IsNullOrEmpty(embeddedLicense))
                rc = true;

            return rc;
        }

        private void Password_Load(object sender, EventArgs e)
        {
            locale = ConfigSetting.Locale;

        }

    }
}
