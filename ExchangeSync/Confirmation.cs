using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Xml;

namespace ExchangeSync
{
    public partial class Confirmation : UserControl
    {
        Guid selectedTemplate = Guid.Empty;

        public Confirmation()
        {
            InitializeComponent();
        }

        private void btnNew_Click(object sender, EventArgs e)
        {
            // Enable Fields
            selectedTemplate = Guid.Empty;
            EnableFields(true);
        }

        private void EnableFields(bool enable)
        {
            txtFrom.Enabled = enable;
            txtTo.Enabled = enable;
            txtCc.Enabled = enable;
            txtMessage.Enabled = enable;
            txtSubject.Enabled = enable;
            txtTemplateName.Enabled = enable;
            // btnNew.Enabled = enable;
            btnSave.Enabled = enable;
            btnTestEmail.Enabled = enable;
            btnSetDefault.Enabled = enable;
        }

        private void Confirmation_Load(object sender, EventArgs e)
        {
            EnableFields(false);
            LoadExistingTemplates();
        }

        private bool ValidateFields()
        {
            bool success = true;
            StringBuilder sb = new StringBuilder();
            sb.AppendLine("The following validation errors occured:"); sb.AppendLine();

            if (!string.IsNullOrEmpty(txtFrom.Text))
            {
                if (!ValidateEmailAddressField(txtFrom.Text))
                {
                    sb.AppendLine("The From Address Field is not properly formatted");
                    success = false;
                }
            }
            else
            {
                sb.AppendLine("The From Address Field is required");
                success = false;
            }

            if (!string.IsNullOrEmpty(txtTo.Text))
            {
                if (!ValidateEmailAddressField(txtTo.Text))
                {
                    sb.AppendLine("The To Address Field is not properly formatted");
                    success = false;
                }
            }
            else
            {
                sb.AppendLine("The To Address Field is required");
                success = false;
            }

            if (!string.IsNullOrEmpty(txtCc.Text))
            {
                if (!ValidateEmailAddressField(txtCc.Text))
                {
                    sb.AppendLine("The Cc Address Field is not properly formatted");
                    success = false;
                }
            }

            if (string.IsNullOrEmpty(txtSubject.Text))
            {
                sb.AppendLine("The Subject Field is required");
                success = false;
            }

            if (string.IsNullOrEmpty(txtMessage.Text))
            {
                sb.AppendLine("The Message Field is required");
                success = false;
            }

            if (!success)
            {
                MessageBox.Show(sb.ToString(), "Validation Errors", MessageBoxButtons.OK);
            }

            return success;
        }

        private string CreateTemplate(string from, string to, string cc, string subject, string body)
        {
            selectedTemplate = Guid.NewGuid();

            StringBuilder sb = new StringBuilder();
            sb.AppendLine(String.Format("<MessageTemplate Id='{0}'>", selectedTemplate.ToString()));
            sb.AppendLine(String.Format("  <From>{0}</From>", from));
            sb.AppendLine(String.Format("  <To>{0}</To>", to));
            sb.AppendLine(String.Format("  <Cc>{0}</Cc>", cc));
            sb.AppendLine(String.Format("  <Subject>{0}</Subject>", subject));
            sb.AppendLine(String.Format("  <Message>{0}</Message>", body));
            sb.AppendLine("</MessageTemplate>");

            return sb.ToString();
        }

        private string UpdateTemplate(Guid templateId, string from, string to, string cc, string subject, string body)
        {
            StringBuilder sb = new StringBuilder();
            sb.AppendLine(String.Format("<MessageTemplate Id='{0}'>", templateId.ToString()));
            sb.AppendLine(String.Format("  <From>{0}</From>", from));
            sb.AppendLine(String.Format("  <To>{0}</To>", to));
            sb.AppendLine(String.Format("  <Cc>{0}</Cc>", cc));
            sb.AppendLine(String.Format("  <Subject>{0}</Subject>", subject));
            sb.AppendLine(String.Format("  <Message>{0}</Message>", body));
            sb.AppendLine("</MessageTemplate>");

            return sb.ToString();
        }

        private bool ValidateEmailAddressField(string fieldValue)
        {
            int bracketPos = 0; string emailAddress = "", displayName = "";
            bool rc = false;

            bracketPos = fieldValue.IndexOf("[");
            if (bracketPos > 0)
            {
                emailAddress = fieldValue.Substring(bracketPos + 1, fieldValue.Length - bracketPos - 2);
                displayName = fieldValue.Substring(0, bracketPos);
            }
            else
            {
                emailAddress = fieldValue;
            }

            if (emailAddress.IsValidEmail())
                rc = true;

            return rc;
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            if (ValidateFields())
            {
                string filename = System.Reflection.Assembly.GetExecutingAssembly().Location;
                string appDir = Path.GetDirectoryName(filename);

                string xmlFileName = appDir + "\\ConfirmationTemplate.xml";

                XmlDocument doc = new XmlDocument();
                doc.Load(xmlFileName);

                XmlNode defaultTemplate = doc.SelectSingleNode("/ConfigurationTemplate/DefaultTemplateId");
                Guid defaultTemplateId = Guid.Empty;
                if (!string.IsNullOrEmpty(defaultTemplate.InnerText))
                    defaultTemplateId = new Guid(defaultTemplate.InnerText);

                if (selectedTemplate != Guid.Empty)
                {
                    // Update Existing Template
                    XmlNodeList nodes = doc.SelectSingleNode("/ConfigurationTemplate/MessageTemplates").ChildNodes;
                    foreach (XmlNode node in nodes)
                    {
                        string nodeIdValue = node.Attributes["Id"].Value;
                        if (!string.IsNullOrEmpty(nodeIdValue))
                        {
                            Guid nodeId = new Guid(nodeIdValue);
                            if (nodeId == defaultTemplateId)
                            {
                                // Update This record
                                XmlNodeList childNodes = node.ChildNodes;
                                foreach (XmlNode childNode in childNodes)
                                {
                                    switch (childNode.Name)
                                    {
                                        case "From":
                                            childNode.InnerText = txtFrom.Text;
                                            break;
                                        case "To":
                                            childNode.InnerText = txtTo.Text;
                                            break;
                                        case "Cc":
                                            childNode.InnerText = txtCc.Text;
                                            break;
                                        case "Subject":
                                            childNode.InnerText = txtSubject.Text;
                                            break;
                                        case "Message":
                                            childNode.InnerText = txtMessage.Text;
                                            break;
                                    }
                                }

                            }
                        }
                    }
                    doc.Save(xmlFileName);
                    ResetForm();
                }
                else
                {
                    // Create New Template
                    XmlNode node = doc.SelectSingleNode("/ConfigurationTemplate/MessageTemplates");
                    
                    XmlNode messageNode = doc.CreateNode(XmlNodeType.Element, "MessageTemplate", null);
                    XmlAttribute messageId = doc.CreateAttribute("Id");
                    XmlAttribute messageName = doc.CreateAttribute("Name");
                    messageId.Value = Guid.NewGuid().ToString();
                    messageName.Value = txtTemplateName.Text;

                    messageNode.Attributes.Append(messageId);
                    messageNode.Attributes.Append(messageName);

                    XmlNode fromNode = doc.CreateNode(XmlNodeType.Element, "From", null);
                    fromNode.InnerText = txtFrom.Text;
                    messageNode.AppendChild(fromNode);

                    XmlNode toNode = doc.CreateNode(XmlNodeType.Element, "To", null);
                    toNode.InnerText = txtTo.Text;
                    messageNode.AppendChild(toNode);

                    XmlNode ccNode = doc.CreateNode(XmlNodeType.Element, "Cc", null);
                    ccNode.InnerText = txtCc.Text;
                    messageNode.AppendChild(ccNode);

                    XmlNode subjectNode = doc.CreateNode(XmlNodeType.Element, "Subject", null);
                    subjectNode.InnerText = txtSubject.Text;
                    messageNode.AppendChild(subjectNode);

                    XmlNode bodyNode = doc.CreateNode(XmlNodeType.Element, "Message", null);
                    bodyNode.InnerText = txtMessage.Text;
                    messageNode.AppendChild(bodyNode);

                    node.AppendChild(messageNode);

                    doc.Save(xmlFileName);
                    ResetForm();
                }
            }
        }

        private void ResetForm()
        {
            txtFrom.Text = string.Empty;
            txtTo.Text = string.Empty;
            txtCc.Text = string.Empty;
            txtSubject.Text = string.Empty;
            txtMessage.Text = string.Empty;
            EnableFields(false);
            btnNew.Enabled = true;
            btnSave.Enabled = false;
            btnTestEmail.Enabled = false;

            LoadExistingTemplates();
        }

        private void LoadExistingTemplates()
        {
            List<KeyValuePair<Guid, string>> list = new List<KeyValuePair<Guid, string>>();

            string filename = System.Reflection.Assembly.GetExecutingAssembly().Location;
            string appDir = Path.GetDirectoryName(filename);

            string xmlFileName = appDir + "\\ConfirmationTemplate.xml";

            XmlDocument doc = new XmlDocument();
            doc.Load(xmlFileName);

            XmlNodeList nodes = doc.SelectSingleNode("/ConfigurationTemplate/MessageTemplates").ChildNodes;
            if (nodes.Count > 0)
            {
                list.Add(new KeyValuePair<Guid, string>(Guid.Empty, "Select a Message Template or Click New"));
                foreach (XmlNode node in nodes)
                {
                    string nodeIdValue = node.Attributes["Id"].Value;
                    string nodeIdText = node.Attributes["Name"].Value;
                    if (!string.IsNullOrEmpty(nodeIdValue))
                    {
                        Guid nodeId = new Guid(nodeIdValue);
                        list.Add(new KeyValuePair<Guid, string>(nodeId, nodeIdText));
                    }
                }
            }

            cmbTemplates.DataSource = list;

        }

        private void btnTestEmail_Click(object sender, EventArgs e)
        {
            if (ValidateFields())
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

                message.CreateEmailMessage(txtFrom.Text, txtTo.Text, txtCc.Text, string.Empty, txtSubject.Text, txtMessage.Text);
                int rc = message.SendEmail();
                message.Dispose();
                
            }
        }

        private void cmbTemplates_SelectedIndexChanged(object sender, EventArgs e)
        {
            int selectedIndex = cmbTemplates.SelectedIndex;
            

            if (selectedIndex > 0)
            {
                string selectedValue = cmbTemplates.SelectedValue.ToString();
                Guid selectedValueId = new Guid(selectedValue);

                selectedTemplate = selectedValueId;
                LoadSelectedTemplate(selectedValueId);
                EnableFields(true);
            }
            else
            {
                ResetForm();
            }

        }

        private void LoadSelectedTemplate(Guid selectedTemplateId)
        {
            string filename = System.Reflection.Assembly.GetExecutingAssembly().Location;
            string appDir = Path.GetDirectoryName(filename);

            string xmlFileName = appDir + "\\ConfirmationTemplate.xml";

            XmlDocument doc = new XmlDocument();
            doc.Load(xmlFileName);

            XmlNodeList nodes = doc.SelectSingleNode("/ConfigurationTemplate/MessageTemplates").ChildNodes;
            foreach (XmlNode node in nodes)
            {
                string nodeIdValue = node.Attributes["Id"].Value;
                if (!string.IsNullOrEmpty(nodeIdValue))
                {
                    Guid nodeId = new Guid(nodeIdValue);
                    if (nodeId == selectedTemplateId)
                    {
                        // Update This record
                        XmlNodeList childNodes = node.ChildNodes;
                        foreach (XmlNode childNode in childNodes)
                        {
                            switch (childNode.Name)
                            {
                                case "From":
                                    txtFrom.Text = childNode.InnerText;
                                    break;
                                case "To":
                                    txtTo.Text = childNode.InnerText;
                                    break;
                                case "Cc":
                                    txtCc.Text = childNode.InnerText;
                                    break;
                                case "Subject":
                                    txtSubject.Text = childNode.InnerText;
                                    break;
                                case "Message":
                                    txtMessage.Text = childNode.InnerText;
                                    break;
                            }
                        }

                    }
                }
            }

        }

        private void btnSetDefault_Click(object sender, EventArgs e)
        {
            int selectedIndex = cmbTemplates.SelectedIndex;
            if (selectedIndex > 0)
            {
                string selectedValue = cmbTemplates.SelectedValue.ToString();

                Core syncCore = ExchangeSyncManager.SyncCore;

                Guid keyId = AppSetting.DefaultConfirmationEmailTemplate.Key;
                AppSetting.DefaultConfirmationEmailTemplate = new KeyValuePair<Guid, string>(keyId, selectedValue);
                syncCore.crm.UpdateApplicationSetting(AppSetting.DefaultConfirmationEmailTemplate.Key, AppSetting.DefaultConfirmationEmailTemplate.GetValue());
            }
            else
            {
                MessageBox.Show("You must select a template in order to set it as default", "Missing Template", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
            }

        }
    }
}
