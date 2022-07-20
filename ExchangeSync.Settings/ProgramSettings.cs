using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExchangeSync
{
    public static class ProgramSetting
    {
        public static string EncryptedSectionName
        {
            get
            {
                return "applicationSettings/ExchangeSync.Properties.Settings";
            }
        }

        public static string CRMSolutionNumber { get; set; }

        public static SourceApplication SourceApp { get; set; }

        /// <summary>
        /// Key contains file name, Value contains profile Display Name
        /// </summary>
        public static KeyValuePair<string, string> ApplicationProfile { get; set; }

        public static List<FieldMap> Map { get; set; }
    }
}
