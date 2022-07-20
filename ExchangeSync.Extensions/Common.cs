using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExchangeSync
{
    public class FieldMap
    {
        public string CRMFieldName { get; set; }
        public string ExchangeFieldName { get; set; }
        public KeyValuePair<int, string> CRMFieldType { get; set; }
    }

    public class ApplicationSetting
    {
        public string Key { get; set; }
        public string Value { get; set; }
        public bool IsSecured { get; set; }
        public string Description { get; set; }

        public ApplicationSetting(string key, string value, string description, bool secured = false)
        {
            Key = key;
            Value = value;
            Description = description;
            IsSecured = secured;
        }
    }

    public class FieldMapping
    {
        public string CRMFieldName { get; set; }

        public string ExchangeFieldName { get; set; }

        public int FieldType { get; set; }

        public int DependencyType { get; set; }

        public FieldMapping(string crmField, string exchangeField, int fieldType, int dependencyType)
        {
            CRMFieldName = crmField;
            ExchangeFieldName = exchangeField;
            FieldType = fieldType;
            DependencyType = dependencyType;
        }
    }

    public class OrganizationalUnit
    {
        public string Name { get; set; }
        public string CanonicalName { get; set; }

        public OrganizationalUnit() { }

        public void Add(string name, string canonicalname)
        {
            this.Name = name;
            this.CanonicalName = canonicalname;
        }
    }

    public class DistributionGroup
    {
        public string Name { get; set; }
        public string OrganizationUnit { get; set; }

        public DistributionGroup() { }

        public void Add(string name, string ou)
        {
            this.Name = name;
            this.OrganizationUnit = ou;
        }
    }

    public class ExchangeContact
    {
        public string Name { get; set; }
        public string Alias { get; set; }
        public string OrganizationalUnit { get; set; }
        public string PrimarySmtpAddress { get; set; }
        public string Identity { get; set; }
    }

    public class SyncEntity
    {
        public string EntityName { get; set; }
        public string EntityDisplayName { get; set; }
        public Guid EntityId { get; set; }
        public string ExchangeGroupName { get; set; }
        public DateTime LastSyncDate { get; set; }
        public bool ExchangeGroupSyncStatus { get; set; }

        public SyncEntity(string entityName, string entityDisplayName, Guid entityId, string exchangeGroupName, DateTime lastSyncDate, bool syncStatus)
        {
            EntityName = entityName;
            EntityDisplayName = entityDisplayName;
            EntityId = entityId;
            ExchangeGroupName = exchangeGroupName;
            LastSyncDate = lastSyncDate;
            ExchangeGroupSyncStatus = syncStatus;
        }
    }

    public class UserJob
    {
        public Guid JobId { get; set; }
        public string JobName { get; set; }
        public string JobType { get; set; }
        public int JobTypeCode { get; set; }
        public string Frequency { get; set; }

        public UserJob() { }
        public UserJob(Guid jobId, string jobName, string jobType, int jobTypeCode, string frequency)
        {
            JobId = jobId;
            JobName = jobName;
            JobType = jobType;
            JobTypeCode = jobTypeCode;
            Frequency = frequency;
        }
    }

    public class AssemblyVersions
    {
        public string AssemblyName { get; set; }
        public string AssemblyVersion { get; set; }
        public string AssemblyOnlineVersion { get; set; }

        public AssemblyVersions(string assemblyName, string assemblyVersion)
        {
            AssemblyName = assemblyName;
            AssemblyVersion = assemblyVersion;
        }
    }

    public class Validator
    {
        public string CRMListName { get; set; }

        public string ExchangeGroupName { get; set; }

        public int TotalCRMContacts { get; set; }

        public int TotalExchangeRecipients { get; set; }

        public Validator(string crmList, string exchangeGroup, int crmContacts, int exchangeRecipients)
        {
            CRMListName = crmList;
            TotalCRMContacts = crmContacts;
            ExchangeGroupName = exchangeGroup;
            TotalExchangeRecipients = exchangeRecipients;
        }
    }

    public static class Helper
    {
        public static string KeyValuePairListToString(List<KeyValuePair<string, string>> list)
        {
            StringBuilder sb = new StringBuilder();

            foreach (KeyValuePair<string, string> pair in list)
            {
                sb.AppendFormat("{0}:{1}; ", pair.Key, pair.Value);
            }

            return sb.ToString();
        }

        public static void LogException(Exception ex, ref string command, ref string function, ref string parameterString)
        {
            string message = ex.Message;
            command = ex.InnerException.Data.Contains("CommandName") ? ex.InnerException.Data["CommandName"].ToString() : string.Empty;
            function = ex.InnerException.Data.Contains("MethodName") ? ex.InnerException.Data["MethodName"].ToString() : string.Empty;

            if (ex.InnerException.Data.Contains("Parameters"))
            {
                List<KeyValuePair<string, string>> parameters = (List<KeyValuePair<string, string>>)ex.InnerException.Data["Parameters"];
                parameterString = Helper.KeyValuePairListToString(parameters);
            }
        }

        public static string CRMServiceUrlToServerUrl(string crmServiceUrl)
        {
            string rc = crmServiceUrl.ToLower();
            // rc = rc.Substring(0, rc.IndexOf("xrmservices") -1 );
            // rc = rc.Substring(rc.LastIndexOf(@"/") + 1);
            return rc;
        }
    }


}
