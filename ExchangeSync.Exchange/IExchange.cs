using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;

namespace ExchangeSync
{
    public enum GroupType
    {
        Distribution = 1,
        Security = 2
    }

    public interface IExchange
    {
        string AliasPrefix { get; set; }

        ExchangeConnectEventArgs Connect();

        ExchangeConnectEventArgs Connect(string exchangeServerUrl, NetworkCredential credentials);

        bool Disconnect();
        void ClearPowerShellCommands();

        string GetServerName();
        bool DistributionGroupExists(string dlname);
        List<string> GetAllDistributionLists();
        List<DistributionGroup> GetAllDistributionListGroups();
        List<OrganizationalUnit> GetAllOrganizationalUnits();
        bool CreateDistributionGroup(string dlname, string ou);
        bool CreateDistributionGroup(string dlname, string ou, GroupType groupType);
        List<ExchangeContact> GetDistributionListMembers(string dlname);
        bool AddDistributionListMember(string dlname, string memberidentity);
        void RemoveAllDistributionListMembers(string dlname);
        bool RemoveDistributionListMember(string dlname, string memberidentity);
        bool ContactExists(string emailaddress);
        ExchangeContact GetMailContact(string identity);
        ExchangeContact GetRecipient(string identity);
        ExchangeContact GetAllMailContacts(string organizationalUnit);
        bool CreateMailContact(string emailaddress, string ou);
        bool CreateMailContact(string displayname, string externalEmailAddress, string firstName, string lastName, string ou, out string alias, string primarySmtpAddress = "");
        bool CreateMailContact(string displayname, string externalEmailAddress, string firstName, string lastName, string ou, string custom, out string alias, string primarySmtpAddress = "");
        bool UserExists(string emailaddress);
        bool ExchangeUserExists(string identity, bool mailboxEnabled);
        bool UpdateMailContact(string identity, string displayname);
        bool UpdateMailContactAlias(string identity, string alias);
        bool UpdateMailContact(string identity, string displayname, string emailaddress);
        bool UpdateMailContact(string identity, string alias, string name, string displayname, string emailaddress);
        bool HideMailContact(string identity, bool hidden);
        bool UpdateMailContact(string identity, string[] parameters, string[] values);
        ExchangeContact GetMailUser(string identity);
        bool UpdateMailUser(string identity, string displayname, string emailaddress);
        bool UpdateMailUser(string identity, string[] parameters, string[] values);
        bool IsOUValid(string ouName);
        void RemoveMailContact(string identity);
        void RemoveDistributionGroup(string groupName);
        void CloseRunspaces();
    }

}
