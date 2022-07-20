using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExchangeSync
{
    public class ItemSync
    {
        public delegate void SyncHandler(ItemSync s, SyncEventArgs e);
        public event SyncHandler Sync;

        public void Start()
        {
            while (true)
            {

            }
        }
    }

    public class SyncEventArgs : EventArgs
    {
        private string emailAddress;

        public string EmailAddress
        {
            get { return emailAddress; }
            set { emailAddress = value; }
        }
    }

    public class Listener
    {
        public void InitiateSync(ItemSync item)
        {
            item.Sync += new ItemSync.SyncHandler(ItemSyncComplete);
        }

        private void ItemSyncComplete(ItemSync item, SyncEventArgs e)
        {
            
        }
    }
}
