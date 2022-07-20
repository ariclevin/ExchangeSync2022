using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExchangeSync
{
    public enum AuthenticationProviderType
    {
        None = 0, // No identity provider. Value = 0.
        ActiveDirectory = 1, // An Active Directory identity provider. Value = 1.
        Federation = 2, // A federated claims identity provider. Value = 2.
        LiveId = 3,  // A pn_Windows_Live_ID identity provider. Value = 3.
        OnlineFederation = 4, // An online (pn_Office_365) federated identity provider. Value = 4.
    }

    public enum ExchangeServerType
    {
        Undefined = 0,
        Exchange_2010 = 2010, // DEPRECATED
        Exchange_2013 = 2013,
        Exchange_2016 = 2016,
        Exchange_2019 = 2019,
        Exchange_2022 = 2022,
        Exchange_Online_365 = 365
    }

    public enum EventLevel
    {
        Verbose = 16,
        Information = 8,
        Warning = 4,
        Error = 2,
        Critical = 1
    }

    public enum PeriodOfDay
    {
        AM = 1,
        PM = 2
    }

    public enum SyncListColumn
    {
        SyncListId = 0,
        ObjectId = 1,
        ObjectTypeName = 2,
        ObjectType = 3,
        ListName = 4,
        ListType = 5,
        ExchangeGroupName = 6,
        GroupType = 7,
        LastSync = 8,
        SyncStatus = 9,
        AutoSyncStatus = 10
    }

    public enum SourceApplication
    {
        Windows = 1,
        Service = 2, // Deprecated
        Console = 3,
        ServiceBus = 4

    }

    public enum AsyncOperationState
    {
        [System.Runtime.Serialization.EnumMemberAttribute()]
        Ready = 0,

        [System.Runtime.Serialization.EnumMemberAttribute()]
        Suspended = 1,

        [System.Runtime.Serialization.EnumMemberAttribute()]
        Locked = 2,

        [System.Runtime.Serialization.EnumMemberAttribute()]
        Completed = 3,
    }
}
