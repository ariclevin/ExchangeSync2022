using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.NetworkInformation;
using System.Text;

namespace ExchangeSync
{
    public static class Network
    {
        public static string MachineName
        {
            get
            {
                return Dns.GetHostEntry(Dns.GetHostName()).HostName;
            }
        }

        public static string HostName
        {
            get
            {
                return Dns.GetHostName();
            }
        }

        public static string DomainName
        {
            get
            {
                var ipProperties = IPGlobalProperties.GetIPGlobalProperties();
                return ipProperties.DomainName; 
            }
        }

        public static string IPAddress
        {
            get
            {
                IPHostEntry ip = Dns.GetHostEntry(Dns.GetHostName());
                IPAddress[] addresses = ip.AddressList;

                string ipAddress = string.Empty;
                foreach (IPAddress address in addresses)
                {
                    ipAddress = address.ToString();
                    if (ipAddress.Contains('.'))
                        break;
                }

                return ipAddress;
            }
        }

        public static string FQDN
        {
            get
            {
                var ipProperties = IPGlobalProperties.GetIPGlobalProperties();
                return string.IsNullOrWhiteSpace(ipProperties.DomainName) ? ipProperties.HostName : string.Format("{0}.{1}", ipProperties.HostName, ipProperties.DomainName);
            }
        }
    }
}
