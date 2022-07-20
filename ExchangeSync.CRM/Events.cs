using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExchangeSync
{
    public class CRMEventArgs : EventArgs
    {
        public enum CRMState
        {
            Undefined = 0,
            Connected = 1,
            NotConnected = 2
        }

        public CRMState ConnectionState { get; set; }

        public string ConnectionReason { get; set; }

        public string LastCRMError { get; set; }

        public bool Result { get; set; }
    }

    [Serializable]
    public class CRMCommandException : Exception
    {
        public CRMCommandException()
            : base()
        { }

        public CRMCommandException(string message)
            : base(message)
        { }

        public CRMCommandException(string format, params object[] args)
            : base(string.Format(format, args))
        { }

        public CRMCommandException(string message, Exception innerException)
            : base(message, innerException)
        {
        }

        public CRMCommandException(string format, Exception innerException, params object[] args)
            : base(string.Format(format, args), innerException)
        { }
    }
}
