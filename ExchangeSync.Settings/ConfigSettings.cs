using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Security;
using System.Text;
using System.Threading.Tasks;

namespace ExchangeSync
{
    // Summary:
    //     Identifies the type of identity provider used for authentication.

    public static class ConfigSetting
    {
        public static bool UseConnectionString { get; set; } // Added 2/5/2019

        public static string CRMServiceUrl { get; set; }
        public static AuthenticationProviderType AuthenticationProvider { get; set; }
        public static bool IntegratedAuthentication { get; set; }
        public static string OrganizationName { get; set; }
        public static string RegionName { get; set; }

        public static string CRMConnectionString { get; set; }

        public static string CRMSolutionVersionNumber { get; set; }

        public static System.Net.NetworkCredential CRMCredentials { get; set; }

        public static string ClientId { get; set; }
        public static string ClientSecret { get; set; }

        public static string ExchangeServerUrl { get; set; }
        public static ExchangeServerType ExchangeServerVersion { get; set; }
        public static System.Net.NetworkCredential ExchangeCredentials { get; set; }

        public static string ContactOU { get; set; }
        public static string DistributionGroupOU { get; set; }

        public static string Locale { get; set;  }

        public static int TimerInterval { get; set; }
    }
}
