using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExchangeSync
{
    public static class AppSetting
    {
        // General Settings
        public static KeyValuePair<Guid, string> ListDeleteAction { get; set; }
        public static KeyValuePair<Guid, string> DuplicateDetectionAction { get; set; }
        public static KeyValuePair<Guid, string> DuplicateSyncAction { get; set; }
        public static KeyValuePair<Guid, string> ExchangeUpdateAction { get; set; }
        public static KeyValuePair<Guid, string> FieldMappingAction { get; set; }
        public static KeyValuePair<Guid, string> HideFromExchangeAddressLists { get; set; }
        public static KeyValuePair<Guid, string> IncludeAutoNumberInAlias { get; set; }


        //CRM Settings
        public static KeyValuePair<Guid, string> ContactRequiredAttributes { get; set; }
        public static KeyValuePair<Guid, string> ContactAutoNumberFieldName { get; set; }
        public static KeyValuePair<Guid, string> ContactRevisionFieldName { get; set; }
        public static KeyValuePair<Guid, string> ContactExchangeAliasFieldName { get; set; }
        public static KeyValuePair<Guid, string> ListDisplayName { get; set; }

        public static KeyValuePair<Guid, string> CompareInactive { get; set; }

        // Exchange Settings
        public static KeyValuePair<Guid, string> ExchangeDistributionGroupOU { get; set; }
        public static KeyValuePair<Guid, string> ExchangeMailContactOU { get; set; }
        public static KeyValuePair<Guid, string> ExchangeMailUserOU { get; set; }

        public static KeyValuePair<Guid, string> ExchangeIdentityType { get; set; }
        public static KeyValuePair<Guid, string> ExchangeAliasFieldFormat { get; set; }
        public static KeyValuePair<Guid, string> ExchangeAliasFieldValue { get; set; }
        public static KeyValuePair<Guid, string> ExchangeNameFieldFormat { get; set; }
        public static KeyValuePair<Guid, string> ExchangeNameFieldValue { get; set; }
        public static KeyValuePair<Guid, string> ExternalDomainName { get; set; }
        public static KeyValuePair<Guid, string> InternalRecipientType { get; set; }

        public static KeyValuePair<Guid, string> ExchangeMailUserUpdate { get; set; }
        public static KeyValuePair<Guid, string> ExchangeMailUserUpdateFields { get; set; }
        public static KeyValuePair<Guid, string> ExchangePolicyAutoUpdateEmailAddresses { get; set; }

        public static KeyValuePair<Guid, string> LastRunStart { get; set; }

        public static KeyValuePair<Guid, string> LastRunEnd { get; set; }

        // Change Tracking Settings
        public static KeyValuePair<Guid, string> DeleteChangeTracking { get; set; }
        public static KeyValuePair<Guid, string> DeleteExchangeContacts { get; set; }

        // Used to determine updating of display names in Exchange (No Updates, CRM Full Name, FirstLast, Last,First)
        public static KeyValuePair<Guid, string> ExchangeMailUserFormat { get; set; }
        public static KeyValuePair<Guid, string> ExchangeMailContactFormat { get; set; }

        public static KeyValuePair<Guid, string> LoggingCriteria { get; set; }
        public static KeyValuePair<Guid, string> LoggingCriteriaVerbose { get; set; }

        public static KeyValuePair<Guid, string> DefaultConfirmationEmailTemplate { get; set; }

        public static KeyValuePair<Guid, string> RestartAutoSyncServicePostOperation { get; set; }
    }

    public class KeyValueSetting
    {
        public Guid Id { get; private set; }
        public string Key { get; private set; }
        public string Value { get; set; }
        public string Description { get; set; }

        public KeyValueSetting(Guid id, string key, string value)
        {
            Id = id;
            Key = key;
            Value = value;
        }

        public KeyValueSetting(Guid id, string key, string value, string description)
        {
            Id = id;
            Key = key;
            Value = value;
            Description = description;
        }
    }
}
