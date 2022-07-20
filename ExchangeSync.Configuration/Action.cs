using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExchangeSync
{
    public partial class StartAction : Form
    {
        public StartAction()
        {
            InitializeComponent();
        }

        private void StartAction_Load(object sender, EventArgs e)
        {
            CheckDefaultConfiguration();
        }

        private void CheckDefaultConfiguration()
        {
            string filename = System.Reflection.Assembly.GetExecutingAssembly().Location;
            string directory = System.IO.Path.GetDirectoryName(filename);

            Configuration config;
            string configFileName = directory + @"\ExchangeSync.exe";

            try
            {
                config = ConfigurationManager.OpenExeConfiguration(configFileName);

                KeyValueConfigurationCollection settings = config.AppSettings.Settings;
                string serviceUrl = settings["CRMServiceUrl"].Value.ToString();

                if (!string.IsNullOrEmpty(serviceUrl))
                {
                    // Has Initial Configuration
                    rbInitialConfiguration.Enabled = true;
                    rbEditProfile.Enabled = true;
                    rbNewProfile.Enabled = true;

                }
                else
                {
                    rbInitialConfiguration.Enabled = true;
                    rbInitialConfiguration.Checked = true;
                    rbEditProfile.Enabled = false;
                    rbNewProfile.Enabled = false;
                }

            }
            catch (System.Configuration.ConfigurationErrorsException ex)
            {
                MessageBox.Show(ex.Message);
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }



        }

        private void btnOK_Click(object sender, EventArgs e)
        {
            Trace.Reset();

            Form nextForm;
            if (rbInitialConfiguration.Checked)
            {
                nextForm = new ConfiguratorWizard();
                nextForm.Show();
                this.Hide();
            }
            else if (rbNewProfile.Checked)
            {
                //nextForm = new ProfileWizard();
                //nextForm.Show();
                //this.Hide();
            }
            else if (rbEditProfile.Checked)
            {
                //if (!string.IsNullOrEmpty(cmbProfileName.Text))
                //{
                //    nextForm = new ProfileWizard(cmbProfileName.Text);
                //   nextForm.Show();
                //    this.Hide();
                //}
                //else
                //{
                //    // Select Profile 
                //}
            }
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            // Exit Application
            DialogResult dr = MessageBox.Show("Are you sure you want to exit the Configuration/Profile Wizard", "Exit Wizard", MessageBoxButtons.YesNo);

            if (dr == DialogResult.Yes)
            {
                Application.Exit();
            }


        }

        private void rbEditProfile_CheckedChanged(object sender, EventArgs e)
        {
            if (rbEditProfile.Checked)
            {
                DataTable dt = new DataTable();
                dt.Columns.Add("Key", Type.GetType("System.String"));
                dt.Columns.Add("Value", Type.GetType("System.String"));

                cmbProfileName.Enabled = true;

                string filename = System.Reflection.Assembly.GetExecutingAssembly().Location;
                string directory = Path.GetDirectoryName(filename);

                string[] files = Directory.GetFiles(directory, "*.config");

                foreach (string fullName in files)
                {
                    string file = fullName.Substring(fullName.LastIndexOf(@"\") + 1);

                    if (file != "ExchangeSync.exe.config")
                    {
                        string profileName = GetConfigFileProfileName(fullName);
                        dt.Rows.Add(file, profileName);
                    }
                }

                cmbProfileName.DataSource = dt;
            }
        }

        private string GetConfigFileProfileName(string fileName)
        {
            Configuration config;
            ExeConfigurationFileMap configFileMap = new ExeConfigurationFileMap();
            configFileMap.ExeConfigFilename = fileName;
            config = ConfigurationManager.OpenMappedExeConfiguration(configFileMap, ConfigurationUserLevel.None);
            KeyValueConfigurationCollection settings = config.AppSettings.Settings;

            string profileName = string.Empty;
            if (settings.AllKeys.Contains<string>("ProfileName"))
            {
                profileName = settings["ProfileName"].Value.ToString();
            }

            return profileName;
        }

        private void rbNewProfile_CheckedChanged(object sender, EventArgs e)
        {
            if (rbNewProfile.Checked)
                cmbProfileName.Enabled = false;
        }

        private void rbInitialConfiguration_CheckedChanged(object sender, EventArgs e)
        {
            if (rbInitialConfiguration.Checked)
                cmbProfileName.Enabled = false;
        }
    }
}
