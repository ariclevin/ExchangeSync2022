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
    public partial class Profile : Form
    {
        public Profile()
        {
            InitializeComponent();
        }

        private void Profile_Load(object sender, EventArgs e)
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("Key", Type.GetType("System.String"));
            dt.Columns.Add("Value", Type.GetType("System.String"));
            
            if (HasDefaultProfile())
            {
                string defaultProfileName = Network.HostName.ToUpper();
                dt.Rows.Add(defaultProfileName, defaultProfileName + " (Default)");                    
            }

            string filename = System.Reflection.Assembly.GetExecutingAssembly().Location;
            string directory = Path.GetDirectoryName(filename);

            string[] files = Directory.GetFiles(directory, "*.config");

            foreach (string fullName in files)
            {
                string file = fullName.Substring(fullName.LastIndexOf(@"\") + 1);
                
                if ((file != "ExchangeSync.exe.config") && (file != "ExchangeSync.vshost.exe.config"))
                {
                    ExeConfigurationFileMap configFileMap = new ExeConfigurationFileMap();
                    configFileMap.ExeConfigFilename = file;

                    Configuration config = ConfigurationManager.OpenMappedExeConfiguration(configFileMap, ConfigurationUserLevel.None);
                    KeyValueConfigurationCollection settings = config.AppSettings.Settings;

                    string profileName = settings["ProfileName"].Value.ToString();

                    dt.Rows.Add(file, profileName);
                }
                    
            }

            cmbProfileName.DataSource = dt;
        }

        private bool HasDefaultProfile()
        {
            bool rc = false;
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
                    rc = true;
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
            return rc;
        }

        private void btnOK_Click(object sender, EventArgs e)
        {
            string key = ((DataRowView)cmbProfileName.SelectedItem).Row[0].ToString();
            string value = string.Empty;
            if (chkAlternateProfile.Checked)
            {
                value = txtAlternateProfile.Text;
            }
            else
            {
                value = ((DataRowView)cmbProfileName.SelectedItem).Row[1].ToString();
            }

            ProgramSetting.ApplicationProfile = new KeyValuePair<string, string>(key, value);
            // ProgramSetting.Map = new List<FieldMap>();

            ExchangeSyncManager nextForm = new ExchangeSyncManager();
            nextForm.Show(this);
            this.Hide();
        }

        private void btnOptions_Click(object sender, EventArgs e)
        {
            this.Height = 205;
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void chkAlternateProfile_CheckedChanged(object sender, EventArgs e)
        {
            if (chkAlternateProfile.Checked)
            {
                txtAlternateProfile.Visible = true;
                cmbProfileName.Visible = false;
            }
            else
            {
                cmbProfileName.Visible = true; 
                txtAlternateProfile.Visible = false;
            }


        }
    }
}
