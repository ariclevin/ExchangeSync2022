using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Management.Automation.Runspaces;

namespace ExchangeSync
{
    public class ExchangeConnectEventArgs : EventArgs
    {
        private RunspaceState state;
        private string reason;

        public RunspaceState ConnectionState { get; set; }

        public string ConnectionReason { get; set; }

        public bool Result { get; set; }
    }

    public class ExchangeCommandEventArgs : EventArgs
    {
        public DateTime EventDateTime { get; set; }

        public string CommandName { get; set; }

        public string ErrorMessage { get; set; }

        public List<KeyValuePair<string, string>> Parameters { get; set; }
    }

    public class ContactCountEventArgs : EventArgs
    {
        public int TotalContacts { get; set; }

        public ContactCountEventArgs(int totalContacts)
        {
            TotalContacts = totalContacts;
        }
    }

    public class ContactSynchedEventArgs : EventArgs
    {
        public string FullName { get; set; }
        public string EmailAddress { get; set; }

        public ContactSynchedEventArgs(string fullName, string emailAddress)
        {
            FullName = fullName;
            EmailAddress = emailAddress;
        }
    }


    [Serializable]
    public class ExchangeCommandException : Exception
    {
        public ExchangeCommandException()
            : base()
        { }

        public ExchangeCommandException(string message)
            : base(message)
        { }

        public ExchangeCommandException(string format, params object[] args)
            : base(string.Format(format, args))
        { }

        public ExchangeCommandException(string message, Exception innerException)
            : base(message, innerException)
        {
        }

        public ExchangeCommandException(string format, Exception innerException, params object[] args)
            : base(string.Format(format, args), innerException)
        { }
    }

}
