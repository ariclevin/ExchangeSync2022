using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Mail;
using System.Text;
using System.Threading.Tasks;

namespace ExchangeSync
{
    public class Messenger
    {

        public string UserName { get; set; }
        public string Password { get; set; }
        public string MailServer { get; set; }
        public int PortNumber { get; set; }
        public string Source { get; set; }

        private MailMessage message;

        /// <summary>
        /// Creates a New Email Message
        /// </summary>
        /// <param name="From">Sender: John Smith [john.smith@contoso.com]</param>
        /// <param name="To">Recipient: John Smith [john.smith@contoso.com]</param>
        /// <param name="Cc">Carbon Copy: John Smith [john.smith@contoso.com]</param>
        /// <param name="Bcc">Blind Carbon Copy: John Smith [john.smith@contoso.com]</param>
        /// <param name="Subject">Subject of the Message</param>
        /// <param name="Body">Content of the Message</param>
        public void CreateEmailMessage(string From, string To, string Cc, string Bcc, string Subject, string Body)
        {
            message = new MailMessage();
            int BracketPos = 0; string eMailAddress = "", DisplayName = "";

            BracketPos = From.IndexOf("[");
            if (BracketPos > 0)
            {
                eMailAddress = From.Substring(BracketPos + 1, From.Length - BracketPos - 2);
                DisplayName = From.Substring(0, BracketPos);
                message.From = new MailAddress(eMailAddress, DisplayName);
            }
            else
                message.From = new MailAddress(From, From);

            BracketPos = To.IndexOf("[");
            if (BracketPos > 0)
            {
                eMailAddress = To.Substring(BracketPos + 1, To.Length - BracketPos - 2);
                DisplayName = To.Substring(0, BracketPos);
                message.To.Add(new MailAddress(eMailAddress, DisplayName));
            }
            else
                message.To.Add(new MailAddress(To, To));

            if (!string.IsNullOrEmpty(Cc))
            {
                BracketPos = Cc.IndexOf("[");
                if (BracketPos > 0)
                {
                    eMailAddress = Cc.Substring(BracketPos + 1, Cc.Length - BracketPos - 2);
                    DisplayName = Cc.Substring(0, BracketPos);
                    message.CC.Add(new MailAddress(eMailAddress, DisplayName));
                }
                else
                    message.CC.Add(new MailAddress(Cc, Cc));
            }
            if (!string.IsNullOrEmpty(Bcc))
            {
                BracketPos = Bcc.IndexOf("[");
                if (BracketPos > 0)
                {
                    eMailAddress = Bcc.Substring(BracketPos + 1, Bcc.Length - BracketPos - 2);
                    DisplayName = Bcc.Substring(0, BracketPos);
                    message.Bcc.Add(new MailAddress(eMailAddress, DisplayName));
                }
                else
                    message.Bcc.Add(new MailAddress(Bcc, Bcc));
            }

            message.Subject = Subject;
            message.Body = Body;
            message.IsBodyHtml = true;
        }

        public int SendEmail()
        {
            SmtpClient client = new SmtpClient(MailServer, 25);
            client.DeliveryMethod = SmtpDeliveryMethod.Network;
            client.Credentials = new System.Net.NetworkCredential(UserName, Password);

            int retVal = 0;
            try
            {
                client.Send(message);
            }
            catch (System.Net.Mail.SmtpException ex)
            {
                Trace.AddLog(EventLevel.Error, DateTime.Now, "Messaging", "SendEmail", "client.Send", ex.Message, "System.Net.Mail.SmtpException", string.Empty);
            }
            catch (System.Exception ex)
            {
                Trace.AddLog(EventLevel.Error, DateTime.Now, "Messaging", "SendEmail", "client.Send", ex.Message, string.Empty, string.Empty);
                // System.Diagnostics.EventLog.WriteEntry("ExchangeSync", ex.Message, System.Diagnostics.EventLogEntryType.Error);
                
            }

            return retVal;
        }

        public void AddAttachment(string fileName)
        {
            Attachment data = new Attachment(fileName);
            message.Attachments.Add(data);
        }

        public void Dispose()
        {
            UserName = string.Empty;
            Password = string.Empty;
            MailServer = string.Empty;
            PortNumber = 0;

            message = null;
        }

    }
}
