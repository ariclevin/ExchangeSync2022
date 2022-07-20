using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Security;
using System.ServiceModel;
using System.ServiceModel.Description;
using System.Text;
using System.Threading.Tasks;

using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Tooling.Connector;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Crm.Sdk.Messages;

namespace ExchangeSync
{
    public enum ProviderType
    {
        ActiveDirectory = 1,
        InternetFacingDeployment = 2,
        OAuth = 3,
        Office365 = 4
    }

    public static class CRMHelper
    {
        /// <summary>
        /// Connect Using Connection String
        /// </summary>
        /// <remarks>Deprecated Function</remarks>
        /// <param name="connectionString"></param>
        /// <returns></returns>
        public static IOrganizationService Connect(string connectionString, string password, bool isEncrypted = false)
        {
            IOrganizationService service;
            if (!string.IsNullOrEmpty(password))
            {
                if (isEncrypted)
                {
                    Security security = new Security();
                    security.CryptoBytes = "BriteOne";
                    string decryptedPassword = security.Decrypt(password);
                    connectionString = connectionString.Replace("CRM_PASSWORD", decryptedPassword);
                }
            }

            CrmServiceClient conn = new CrmServiceClient(connectionString);
            if (conn.IsReady)
            {
                service = (IOrganizationService)conn.OrganizationWebProxyClient != null ? (IOrganizationService)conn.OrganizationWebProxyClient : (IOrganizationService)conn.OrganizationServiceProxy;
                return service;
            }
            else
            {
                return null;
            }

        }

        /// <summary>
        /// Connect to AD
        /// </summary>
        /// <param name="userName">domain\username</param>
        /// <param name="password">password</param>
        /// <param name="internalUrl">servername.domain.com</param>
        /// <param name="organizationName">orgname</param>
        /// <param name="isEncrypted">is Password Encrypted By DESCryptoServiceProvider</param>
        /// <returns></returns>
        public static CrmServiceClient Connect(string userName, string password, string domain, string server, int port, string organizationName, bool useSSL = false, bool isEncrypted = false)
        {
            IOrganizationService service;
            if (!string.IsNullOrEmpty(password))
            {
                if (isEncrypted)
                {
                    Security security = new Security();
                    security.CryptoBytes = "BriteOne";
                    password = security.Decrypt(password);
                }
            }

            NetworkCredential creds = new NetworkCredential(userName, password, domain);
            Microsoft.Xrm.Tooling.Connector.AuthenticationType authType = Microsoft.Xrm.Tooling.Connector.AuthenticationType.AD;
            CrmServiceClient conn = new Microsoft.Xrm.Tooling.Connector.CrmServiceClient(creds, authType, server, port.ToString(), organizationName, false, useSSL, null);

            return conn; 
        }


        /// <summary>
        /// Connect to IFD
        /// </summary>
        /// <param name="userName">domain\username</param>
        /// <param name="password">password</param>
        /// <param name="internalUrl">servername.domain.com</param>
        /// <param name="organizationName">orgname</param>
        /// <param name="isEncrypted">is Password Encrypted By DESCryptoServiceProvider</param>
        /// <returns></returns>
        public static CrmServiceClient Connect(string userName, string password, string internalUrl, string organizationName, bool isEncrypted = false)
        {
            IOrganizationService service;
            if (!string.IsNullOrEmpty(password))
            {
                if (isEncrypted)
                {
                    Security security = new Security();
                    security.CryptoBytes = "BriteOne";
                    password = security.Decrypt(password);
                }
            }

            NetworkCredential creds = new NetworkCredential(userName, ConvertToSecureString(password));
            Microsoft.Xrm.Tooling.Connector.AuthenticationType authType = Microsoft.Xrm.Tooling.Connector.AuthenticationType.IFD;
            CrmServiceClient conn = new Microsoft.Xrm.Tooling.Connector.CrmServiceClient(creds, authType, internalUrl, "443", organizationName, true, true, null);
            return conn;
        }

        /// <summary>
        /// Connect to Office 365
        /// </summary>
        /// <param name="userName">Office 365 Username</param>
        /// <param name="password">Office 365 Password</param>
        /// <param name="orgName">Office 365 Organization Name</param>
        /// <param name="region">Office 365 Region</param>
        /// <returns></returns>
        public static CrmServiceClient Connect(string userName, string password, string orgName, string region = "NorthAmerica")
        {
            IOrganizationService service;
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
            CrmServiceClient conn = new CrmServiceClient(userName, ConvertToSecureString(password), region, orgName, isOffice365: true);
            
            return conn;
        }

        public static CrmServiceClient Connect(string connectionString)
        {
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
            CrmServiceClient conn = new CrmServiceClient(connectionString);
            return conn;
        }

        public static CrmServiceClient Connect(string serverUrl, string clientId, string clientSecret, bool useConnectionString = true)
        {
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
            string connectionString = String.Format("AuthType=ClientSecret;url={0};ClientId={1};ClientSecret={2}", serverUrl, clientId, clientSecret);
            CrmServiceClient conn = new CrmServiceClient(connectionString);
            
            return conn;
        }

        private static System.Security.SecureString ConvertToSecureString(string password)
        {
            if (password == null)
                throw new ArgumentNullException("Missing Password");

            var securePassword = new System.Security.SecureString();
            foreach (char c in password)
                securePassword.AppendChar(c);

            securePassword.MakeReadOnly();
            return securePassword;
        }

    }
}
