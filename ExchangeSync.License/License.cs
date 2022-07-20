using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExchangeSync
{
    public static class License
    {
        public static bool WindowsForms { get; set; }
        public static bool AutoSync { get; set; }
        // public static bool ConsoleSync { get; set; }

        public static bool Evaluation { get; set; }

        public static bool EvaluationExpired { get; set; }

        public static string CompanyName { get; set; }

        public static string CompanyWebSite { get; set; }

        public static string RegisteredToFirstName { get; set; }

        public static string RegisteredToLastName { get; set; }

        public static string RegisteredToPhoneNumber { get; set; }

        public static string RegisteredToEmailAddress { get; set; }

        public static string MaintenanceExpirationDate { get; set; }

        public static string LicenseKey { get; set; }

        // Options are: Evaluation, Community, Enterprise
        public static string LicenseType { get; set; }

        public static DateTime ExpirationDate { get; set; }

        public static string ComputerName { get; set; }


    }
}
