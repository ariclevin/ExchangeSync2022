using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Client;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Crm.Sdk.Messages;

namespace ExchangeSync
{
    public partial class Core
    {
        public IExchange exch;
        public CRMService crm;

        public bool IsCRMConnected { get; set; }
        public bool IsExchangeConnected { get; set; }

        public delegate void ContactSynchedHandler(object sender, ContactSynchedEventArgs e);
        public delegate void ContactCountHandler(object sender, ContactCountEventArgs e);
        public event ContactSynchedHandler ContactSynched;
        public event ContactCountHandler ContactsCounted;

        #region Application and Configuration Settings

        public bool LoadLicensingInformation()
        {
            string filename = System.Reflection.Assembly.GetExecutingAssembly().Location;
            string directory = System.IO.Path.GetDirectoryName(filename);

            string licenseFile = directory + @"\ExchangeSync.lic";

            bool licenseExists = File.Exists(licenseFile);

            if (licenseExists)
            {
                string license = File.ReadAllText(licenseFile);

                ExchangeSync.CoreLicense coreLicense = new ExchangeSync.CoreLicense(license);
                coreLicense.InitLicense();
            }

            return licenseExists;
        }


        public void LoadConfigurationSettings()
        {
            string filename = System.Reflection.Assembly.GetExecutingAssembly().Location;
            string directory = System.IO.Path.GetDirectoryName(filename);

            string configFileName = directory + @"\ExchangeSync.exe";
            Configuration config = ConfigurationManager.OpenExeConfiguration(configFileName);
            
            KeyValueConfigurationCollection settings = config.AppSettings.Settings;

            ConfigSetting.CRMServiceUrl = settings["CRMServiceUrl"].Value.ToString();
            if (!string.IsNullOrEmpty(settings["CRMAuthenticationProviderType"].Value))
            {
                ConfigSetting.AuthenticationProvider = (AuthenticationProviderType)Convert.ToInt32(settings["CRMAuthenticationProviderType"].Value);
            }
            else
            {
                ConfigSetting.AuthenticationProvider = AuthenticationProviderType.None;
            }

            if (!string.IsNullOrEmpty(settings["CRMIntegratedAuthentication"].Value))
            {
                ConfigSetting.IntegratedAuthentication = Convert.ToBoolean(Convert.ToInt32(settings["CRMIntegratedAuthentication"].Value));
            }
            else
            {
                ConfigSetting.IntegratedAuthentication = false;
            }

            if (!string.IsNullOrEmpty(settings["UseCRMConnectionString"].Value))
            {
                ConfigSetting.UseConnectionString = Convert.ToBoolean(Convert.ToInt32(settings["UseCRMConnectionString"].Value));
            }
            else
            {
                ConfigSetting.UseConnectionString = false;
            }


            System.Net.NetworkCredential crmCredential = new System.Net.NetworkCredential();
            crmCredential.UserName = settings["CRMUsername"].Value;
            crmCredential.Domain = settings["CRMDomain"].Value;

            ConfigSetting.CRMCredentials = crmCredential;
            ConfigSetting.OrganizationName = settings["CRMOrganization"].Value;
            ConfigSetting.RegionName = settings["CRMRegion"].Value;

            ConfigSetting.ClientId = settings["ClientId"].Value;
            ConfigSetting.ClientSecret = settings["ClientSecret"].Value;


            ConfigSetting.ExchangeServerUrl = settings["ExchangeServerName"].Value;

            if (!string.IsNullOrEmpty(settings["ExchangeServerVersion"].Value))
            {
                int exchangeServerVersion = 0;
                bool isInt = int.TryParse(settings["ExchangeServerVersion"].Value, out exchangeServerVersion);
                if (isInt)
                {
                    ConfigSetting.ExchangeServerVersion = (ExchangeServerType)exchangeServerVersion;
                }
                else
                {
                    if (settings["ExchangeServerVersion"].Value == "Online")
                        ConfigSetting.ExchangeServerVersion = ExchangeServerType.Exchange_Online_365;
                    else
                        ConfigSetting.ExchangeServerVersion = ExchangeServerType.Undefined;
                }
                
            }
            else
            {
                ConfigSetting.ExchangeServerVersion = ExchangeServerType.Undefined;
            }

            System.Net.NetworkCredential adCredential = new System.Net.NetworkCredential();
            adCredential.UserName = settings["ADUsername"].Value;
            adCredential.Domain = settings["ADDomain"].Value;

            ConfigSetting.ExchangeCredentials = adCredential;

            ConnectionStringsSection csSection = config.ConnectionStrings;
            ConnectionStringSettingsCollection csCollection = csSection.ConnectionStrings;
            ConnectionStringSettings cs = csCollection["Dynamics"];
            ConfigSetting.CRMConnectionString = cs.ConnectionString;

            // Need to get encrypted passwords
            ConfigurationSection section = config.GetSection(ProgramSetting.EncryptedSectionName);
            ClientSettingsSection clientSection = (ClientSettingsSection)section;

            SettingElement seCRMPassword = clientSection.Settings.Get("CRMPassword");
            crmCredential.Password = seCRMPassword.Value.ValueXml.InnerXml;
            ConfigSetting.CRMCredentials = crmCredential;

            SettingElement seADPassword = clientSection.Settings.Get("ADPassword");
            adCredential.Password = seADPassword.Value.ValueXml.InnerXml;
            ConfigSetting.ExchangeCredentials = adCredential;

            ConfigSetting.TimerInterval = Convert.ToInt32(settings["TimerInterval"].Value);
            ConfigSetting.Locale = settings["Locale"].Value;
        }

        public void LoadConfigurationSettings(string profileName)
        {
            string filename = System.Reflection.Assembly.GetExecutingAssembly().Location;
            string directory = System.IO.Path.GetDirectoryName(filename);

            ExeConfigurationFileMap configFileMap = new ExeConfigurationFileMap();
            configFileMap.ExeConfigFilename = profileName;

            Configuration config = ConfigurationManager.OpenMappedExeConfiguration(configFileMap, ConfigurationUserLevel.None);
            KeyValueConfigurationCollection settings = config.AppSettings.Settings;

            ConfigSetting.CRMServiceUrl = settings["CRMServiceUrl"].Value.ToString();
            if (!string.IsNullOrEmpty(settings["CRMAuthenticationProviderType"].Value))
            {
                ConfigSetting.AuthenticationProvider = (AuthenticationProviderType)Convert.ToInt32(settings["CRMAuthenticationProviderType"].Value);
            }
            else
            {
                ConfigSetting.AuthenticationProvider = AuthenticationProviderType.None;
            }

            System.Net.NetworkCredential crmCredential = new System.Net.NetworkCredential();
            crmCredential.UserName = settings["CRMUsername"].Value;
            crmCredential.Domain = settings["CRMDomain"].Value;

            ConfigSetting.ClientId = settings["ClientId"].Value;
            ConfigSetting.ClientSecret = settings["ClientSecret"].Value;

            ConfigSetting.CRMCredentials = crmCredential;

            ConfigSetting.ExchangeServerUrl = settings["ExchangeServerName"].Value;

            if (!string.IsNullOrEmpty(settings["ExchangeServerVersion"].Value))
            {
                ConfigSetting.ExchangeServerVersion = (ExchangeServerType)Convert.ToInt32(settings["ExchangeServerVersion"].Value);
            }
            else
            {
                ConfigSetting.ExchangeServerVersion = ExchangeServerType.Undefined;
            }

            System.Net.NetworkCredential adCredential = new System.Net.NetworkCredential();
            adCredential.UserName = settings["ADUsername"].Value;
            adCredential.Domain = settings["ADDomain"].Value;
            ConfigSetting.ExchangeCredentials = adCredential;

            ConnectionStringsSection csSection = config.ConnectionStrings;
            ConnectionStringSettingsCollection csCollection = csSection.ConnectionStrings;
            ConnectionStringSettings cs = csCollection["Dynamics"];
            ConfigSetting.CRMConnectionString = cs.ConnectionString;

            // Need to get encrypted passwords
            ConfigurationSection section = config.GetSection(ProgramSetting.EncryptedSectionName);
            ClientSettingsSection clientSection = (ClientSettingsSection)section;

            SettingElement seCRMPassword = clientSection.Settings.Get("CRMPassword");
            crmCredential.Password = seCRMPassword.Value.ValueXml.InnerXml;
            ConfigSetting.CRMCredentials = crmCredential;

            SettingElement seADPassword = clientSection.Settings.Get("ADPassword");
            adCredential.Password = seADPassword.Value.ValueXml.InnerXml;
            ConfigSetting.ExchangeCredentials = adCredential;
        }

        public void LoadApplicationSettings()
        {
            EntityCollection settings = new EntityCollection();
            
            try
            {
                settings = crm.RetrieveApplicationSettings(11);
            }
            catch (System.Exception ex)
            {
                string message = ex.Message;
                string command = ex.InnerException.Data.Contains("CommandName") ? ex.InnerException.Data["CommandName"].ToString() : string.Empty;
                string function = ex.InnerException.Data.Contains("MethodName") ? ex.InnerException.Data["MethodName"].ToString() : string.Empty;

                List<KeyValuePair<string, string>> parameters = ex.InnerException.Data.Contains("Parameters") ? (List<KeyValuePair<string, string>>)ex.InnerException.Data["Parameters"] : new List<KeyValuePair<string, string>>();
                string parameterString = Helper.KeyValuePairListToString(parameters);
                Trace.AddLog(EventLevel.Error, DateTime.Now, "CRM", function, command, message, string.Empty, parameterString);
            }

            if (settings.Entities.Count > 0)
            {
                AppSetting.LastRunStart = new KeyValuePair<Guid, string>(Guid.Empty, "");
                AppSetting.LastRunEnd = new KeyValuePair<Guid, string>(Guid.Empty, "");

                foreach (Entity setting in settings.Entities)
                {
                    Guid id = setting.Id;
                    string key = setting.Attributes["xrm_key"].ToString();
                    string value = setting.Contains("xrm_value") ? setting.Attributes["xrm_value"].ToString() : "";
                    
                    switch (key.ToUpper())
                    {
                        case "LISTDELETEACTION":
                            AppSetting.ListDeleteAction = new KeyValuePair<Guid, string>(id, value);
                            break;
                        case "DUPLICATEDETECTIONACTION":
                            AppSetting.DuplicateDetectionAction = new KeyValuePair<Guid, string>(id, value);
                            break;
                        case "DUPLICATESYNCACTION":
                            AppSetting.DuplicateSyncAction = new KeyValuePair<Guid, string>(id, value);
                            break;
                        case "EXCHANGEUPDATEACTION":
                            AppSetting.ExchangeUpdateAction = new KeyValuePair<Guid, string>(id, value);
                            break;
                        case "FIELDMAPPINGACTION":
                            AppSetting.FieldMappingAction = new KeyValuePair<Guid, string>(id, value);
                            break;
                        case "HIDEFROMEXCHANGEADDRESSLISTS":
                            AppSetting.HideFromExchangeAddressLists = new KeyValuePair<Guid, string>(id, value);
                            break;
                        case "CONTACTAUTONUMBERFIELDNAME":
                            AppSetting.ContactAutoNumberFieldName = new KeyValuePair<Guid, string>(id, value);
                            break;
                        case "CONTACTREVISIONFIELDNAME":
                            AppSetting.ContactRevisionFieldName = new KeyValuePair<Guid, string>(id, value);
                            break;
                        case "EXCHANGEALIASFIELDNAME":
                        case "CONTACTEXCHANGEALIASFIELDNAME":
                            AppSetting.ContactExchangeAliasFieldName = new KeyValuePair<Guid, string>(id, value);
                            break;
                        case "CONTACTREQUIREDFIELDS":
                        case "CONTACTREQUIREDATTRIBUTES":
                            AppSetting.ContactRequiredAttributes = new KeyValuePair<Guid, string>(id, value);
                            break;
                        case "LISTDISPLAYNAME":
                            AppSetting.ListDisplayName = new KeyValuePair<Guid, string>(id, value);
                            break;
                        case "COMPAREINACTIVE":
                            AppSetting.CompareInactive = new KeyValuePair<Guid, string>(id, value);
                            break;
                        case "EXCHANGEDISTRIBUTIONGROUPOU":
                            AppSetting.ExchangeDistributionGroupOU = new KeyValuePair<Guid, string>(id, value);
                            break;
                        case "EXCHANGEMAILCONTACTOU":
                            AppSetting.ExchangeMailContactOU = new KeyValuePair<Guid, string>(id, value);
                            break;
                        case "EXCHANGEMAILUSEROU":
                            AppSetting.ExchangeMailUserOU = new KeyValuePair<Guid, string>(id, value);
                            break;
                        case "EXTERNALDOMAINNAME":
                            AppSetting.ExternalDomainName = new KeyValuePair<Guid, string>(id, value);
                            break;
                        case "ALIASFIELDFORMAT":
                        case "EXCHANGEALIASFIELDFORMAT":
                            AppSetting.ExchangeAliasFieldFormat = new KeyValuePair<Guid, string>(id, value);
                            break;
                        case "ALIASFIELDVALUE":
                        case "EXCHANGEALIASFIELDVALUE":
                            AppSetting.ExchangeAliasFieldValue = new KeyValuePair<Guid, string>(id, value);
                            break;
                        case "NAMEFIELDFORMAT":
                        case "EXCHANGENAMEFIELDFORMAT":
                            AppSetting.ExchangeNameFieldFormat = new KeyValuePair<Guid, string>(id, value);
                            break;
                        case "NAMEFIELDVALUE":
                        case "EXCHANGENAMEFIELDVALUE":
                            AppSetting.ExchangeNameFieldValue = new KeyValuePair<Guid, string>(id, value);
                            break;
                        case "EXCHANGEMAILUSERUPDATE":
                            AppSetting.ExchangeMailUserUpdate = new KeyValuePair<Guid, string>(id, value);
                            break;
                        case "EXCHANGEMAILUSERUPDATEFIELDS":
                            AppSetting.ExchangeMailUserUpdateFields = new KeyValuePair<Guid, string>(id, value);
                            break;
                        case "EXCHANGEPOLICYAUTOUPDATEEMAILADDRESSES":
                            AppSetting.ExchangePolicyAutoUpdateEmailAddresses = new KeyValuePair<Guid, string>(id, value);
                            break;
                        case "EXCHANGEMAILUSERFORMAT":
                            AppSetting.ExchangeMailUserFormat = new KeyValuePair<Guid, string>(id, value);
                            break;
                        case "EXCHANGEMAILCONTACTFORMAT":
                            AppSetting.ExchangeMailContactFormat = new KeyValuePair<Guid, string>(id, value);
                            break;
                        case "LOGGINGCRITERIA":
                            AppSetting.LoggingCriteria = new KeyValuePair<Guid, string>(id, value);
                            break;
                        case "LOGGINGCRITERIAVERBOSE":
                            AppSetting.LoggingCriteriaVerbose = new KeyValuePair<Guid, string>(id, value);
                            break;
                        case "INTERNALRECIPIENTTYPE":
                            AppSetting.InternalRecipientType = new KeyValuePair<Guid, string>(id, value);
                            break;
                        case "EXCHANGEIDENTITYTYPE":
                            AppSetting.ExchangeIdentityType = new KeyValuePair<Guid, string>(id, value);
                            break;
                        case "DEFAULTCONFIRMATIONEMAILTEMPLATE":
                            AppSetting.DefaultConfirmationEmailTemplate = new KeyValuePair<Guid, string>(id, value);
                            break;
                        case "RESTARTAUTOSYNCSERVICEPOSTOPERATION": 
                            AppSetting.RestartAutoSyncServicePostOperation = new KeyValuePair<Guid, string>(id, value);
                            break;
                        case "DELETEEXCHANGECONTACTS":
                            AppSetting.DeleteExchangeContacts = new KeyValuePair<Guid, string>(id, value);
                            break;
                        case "DELETECHANGETRACKING":
                            AppSetting.DeleteChangeTracking = new KeyValuePair<Guid, string>(id, value);
                            break;
                        case "LASTRUNSTART":
                            AppSetting.LastRunStart = new KeyValuePair<Guid, string>(id, value);
                            break;
                        case "LASTRUNEND":
                            AppSetting.LastRunEnd = new KeyValuePair<Guid, string>(id, value);
                            break;
                    }
                }
            }

        }

        public void UpdateSyncStart()
        {
            DateTime lastRunStart = DateTime.Now;
            if (AppSetting.LastRunStart.Key != Guid.Empty)
                crm.UpdateApplicationSetting(AppSetting.LastRunStart.Key, lastRunStart.ToString());
            else
            {
                Guid lastRunStartId = crm.CreateApplicationSetting(11, "LASTRUNSTART", lastRunStart.ToString(), "The Date and Time that the last synchronization process initiated");
                AppSetting.LastRunStart = new KeyValuePair<Guid, string>(lastRunStartId, lastRunStart.ToString());
            }
                
        }

        public void UpdateSyncEnd()
        {
            DateTime lastRunEnd = DateTime.Now;
            if (AppSetting.LastRunEnd.Key != Guid.Empty)
                crm.UpdateApplicationSetting(AppSetting.LastRunEnd.Key, lastRunEnd.ToString());
            else
            {
                Guid lastRunEndId = crm.CreateApplicationSetting(11, "LASTRUNEND", lastRunEnd.ToString(), "The Date and Time that the last synchronization process completed");
                AppSetting.LastRunEnd = new KeyValuePair<Guid, string>(lastRunEndId, lastRunEnd.ToString());
            }
        }


        public List<FieldMap> LoadDefaultColumnSet()
        {
            List<FieldMap> map = new List<FieldMap>();
            EntityCollection fields = new EntityCollection();

            try
            {
                fields = crm.RetrieveFieldMap();
            }
            catch (System.Exception ex)
            {
                string message = ex.Message;
                string command = ex.InnerException.Data.Contains("CommandName") ? ex.InnerException.Data["CommandName"].ToString() : string.Empty;
                string function = ex.InnerException.Data.Contains("MethodName") ? ex.InnerException.Data["MethodName"].ToString() : string.Empty;

                List<KeyValuePair<string, string>> parameters = ex.InnerException.Data.Contains("Parameters") ? (List<KeyValuePair<string, string>>)ex.InnerException.Data["Parameters"] : new List<KeyValuePair<string, string>>();
                string parameterString = Helper.KeyValuePairListToString(parameters);
                Trace.AddLog(EventLevel.Error, DateTime.Now, "CRM", function, command, message, string.Empty, parameterString);
            }
            
            if (fields.Entities.Count > 0)
            {
                foreach (Entity field in fields.Entities)
                {
                    string crmFieldName = field.Attributes["xrm_fieldname"].ToString();
                    int crmFieldType = ((OptionSetValue)(field.Attributes["xrm_type"])).Value;
                    string crmFieldTypeName = field.FormattedValues["xrm_type"].ToString();
                    string exchangeFieldName = field.Attributes["xrm_exchangefieldname"].ToString().Trim();

                    FieldMap fieldMapping = new FieldMap();
                    fieldMapping.CRMFieldName = crmFieldName;
                    fieldMapping.ExchangeFieldName = exchangeFieldName;
                    fieldMapping.CRMFieldType = new KeyValuePair<int, string>(crmFieldType, crmFieldTypeName);
                    map.Add(fieldMapping);
                }
            }
            return map;
        }

        #endregion

        #region CRM and Exchange Connection Methods

        public ExchangeConnectEventArgs ConnectExchange()
        {
            switch (ConfigSetting.ExchangeServerVersion)
            {
                // Not Required. Handled in Default
                case ExchangeServerType.Exchange_2010:
                case ExchangeServerType.Exchange_2013:
                case ExchangeServerType.Exchange_2016:
                case ExchangeServerType.Exchange_2019:
                case ExchangeServerType.Exchange_2022:
                    exch = new ExchangeBase();
                    break;
                case ExchangeServerType.Exchange_Online_365:
                    exch = new Exchange365();
                    break;
                default:
                    exch = new ExchangeBase();
                    break;
            }

            ExchangeConnectEventArgs args = exch.Connect();
            IsExchangeConnected = args.Result;
            return args;
        }

        public ExchangeConnectEventArgs ConnectExchange(string exchangeServerUrl, ExchangeServerType versionType, NetworkCredential credentials)
        {
            if (versionType != ExchangeServerType.Undefined)
            {
                switch (versionType)
                {
                    case ExchangeServerType.Exchange_2010:
                    case ExchangeServerType.Exchange_2013:
                    case ExchangeServerType.Exchange_2016:
                    case ExchangeServerType.Exchange_2019:
                    case ExchangeServerType.Exchange_2022:
                        exch = new ExchangeBase();
                        break;
                    case ExchangeServerType.Exchange_Online_365:
                        exch = new Exchange365();
                        break;
                    default:
                        exch = new ExchangeBase();
                        break;
                }
            }

            ExchangeConnectEventArgs args = exch.Connect(exchangeServerUrl, credentials);
            IsExchangeConnected = args.Result;
            return args;

        }

        public CRMEventArgs ConnectToCRM(int providerType, NetworkCredential credentials, string serverName, string organizationName, string regionName)
        {
            crm = new CRMService();
            CRMEventArgs args = crm.Connect(providerType, credentials, serverName, organizationName, regionName);
            IsCRMConnected = (args.ConnectionState == CRMEventArgs.CRMState.Connected) ? true : false;
            return args;
        }

        public CRMEventArgs ConnectToCRM(string connectionString, string password)
        {
            crm = new CRMService();
            connectionString = connectionString.Replace("********", password);
            CRMEventArgs args = crm.Connect(connectionString);
            IsCRMConnected = (args.ConnectionState == CRMEventArgs.CRMState.Connected) ? true : false;
            return args;
        }

        public CRMEventArgs ConnectToCRM(string crmserverUrl, string clientId, string clientSecret)
        {
            crm = new CRMService();
            CRMEventArgs args = crm.Connect(crmserverUrl, clientId, clientSecret);
            IsCRMConnected = (args.ConnectionState == CRMEventArgs.CRMState.Connected) ? true : false;
            return args;
        }


        #endregion

        #region Exchange Contact and User Update Methods

        public void ExecuteChangeTrackingOperations()
        {
            if (AppSetting.DeleteExchangeContacts.Value.ToUpper() == "YES")
            {
                EntityCollection operations = crm.RetrieveChangeTrackingOperations(2);
                if (operations.Entities.Count > 0)
                {
                    foreach (Entity cto in operations.Entities)
                    {
                        Guid operationId = cto.Id;
                        string alias = cto.Contains("xrm_alias") ? cto["xrm_alias"].ToString() : string.Empty;
                        string emailAddress = cto.Contains("xrm_emailaddress") ? cto["xrm_emailaddress"].ToString() : string.Empty;

                        bool contactFound = false;
                        if (!string.IsNullOrEmpty(emailAddress)) // Remove by Email Address
                        {
                            if (ContactExists(emailAddress))
                            {
                                contactFound = true;
                                try
                                {
                                    exch.RemoveMailContact(emailAddress);
                                    Trace.AddLog(EventLevel.Information, DateTime.Now, "ExecuteChangeTrackingOperations", "Exchange", "Removing Exchange Email Address: " + emailAddress, "");
                                }
                                catch (System.Exception ex)
                                {
                                    Trace.AddLog(EventLevel.Error, DateTime.Now, "ExecuteChangeTrackingOperations", "Exchange", "Error Removing Exchange Email Address: " + emailAddress, ex.Message);
                                }
                            }
                        }

                        if (!contactFound) // Remove by Alias
                        {
                            if (!string.IsNullOrEmpty(alias))
                            {
                                if (ContactExists(alias))
                                {
                                    contactFound = true;
                                    try
                                    {
                                        exch.RemoveMailContact(alias);
                                        Trace.AddLog(EventLevel.Information, DateTime.Now, "ExecuteChangeTrackingOperations", "Exchange", "Removing Exchange Alias: " + alias, "");
                                    }
                                    catch (System.Exception ex)
                                    {
                                        Trace.AddLog(EventLevel.Error, DateTime.Now, "ExecuteChangeTrackingOperations", "Exchange", "Error Removing Exchange Alias: " + alias, ex.Message);
                                    }
                                }
                            }
                        }

                        if (!contactFound)
                        {
                            Trace.AddLog(EventLevel.Warning, DateTime.Now, "ExecuteChangeTrackingOperations", "Exchange", "No Alias or Email Address exist for Change Tracking Operation for Record Id: " + operationId.ToString(), "");
                        }
                        else
                        {
                            if (AppSetting.DeleteChangeTracking.Value.ToUpper() == "YES")
                                crm.DeleteChangeTrackOperation(operationId, alias, emailAddress);
                            else
                                crm.UpdateChangeTrackingOperation(operationId, alias, emailAddress);
                        }

                    }
                }
            }
        }

        public void ExecuteSync(Guid syncListId, Guid listId, string entityName, string listName, string exchangeListName, DateTime lastSync, string listType = "")
        {
            switch (listType)
            {
                case "Account":
                    break;
                case "Contact":
                    ExecuteContactSync(syncListId, listId, entityName, listName, exchangeListName, lastSync);
                    break;
                case "Lead":
                    break;
                default: // Contact
                    ExecuteContactSync(syncListId, listId, entityName, listName, exchangeListName, lastSync);
                    break;
            }
        }

        private void ExecuteContactSync(Guid syncListId, Guid listId, string entityName, string listName, string exchangeListName, DateTime lastSync)
        {
            bool stopListProcessing = false;
            EntityCollection contacts = RetrieveContacts(entityName, listName, listId);
            if (contacts.Entities.Count > 0)
            {
                if (ProgramSetting.SourceApp == SourceApplication.Windows)
                    ContactsCounted(this, new ContactCountEventArgs(contacts.Entities.Count));

                Trace.AddLog(EventLevel.Information, DateTime.Now, "RunSync", "CRM", "Total  " + contacts.Entities.Count.ToString() + " CRM Contacts retrieved", string.Empty);

                if (AppSetting.ExchangeUpdateAction.Value.ToUpper() == "ALL")
                {
                    try
                    {
                        exch.RemoveAllDistributionListMembers(exchangeListName);
                    }
                    catch (System.Exception ex)
                    {
                        Trace.AddLog(EventLevel.Error, DateTime.Now, "ExecuteContactSync", "Exchange", ex.Message, ex.StackTrace);
                        stopListProcessing = true;
                    }
                }

                string customFieldName = AppSetting.ContactAutoNumberFieldName.Value;
                if (!stopListProcessing)
                {
                    if (AppSetting.FieldMappingAction.Value == "CRM")
                    {
                        string alias = "";
                        Trace.AddLog(EventLevel.Verbose, DateTime.Now, "RunSync", "CRM", "Using CRM Field Mapping", "");
                        UpdateExchangeContacts(exchangeListName, AppSetting.ExternalDomainName.Value, ref alias, ref contacts);
                    }
                    else
                    {
                        if (customFieldName.Contains(':'))
                            customFieldName = customFieldName.Substring(0, customFieldName.IndexOf(':'));
                        Trace.AddLog(EventLevel.Verbose, DateTime.Now, "RunSync", "CRM", String.Format("Custom Field Name = {0}", customFieldName), "");

                        foreach (Entity contact in contacts.Entities)
                        {
                            Guid contactid = contact.Id;
                            string firstname = contact.Attributes.ContainsKey("firstname") ? contact.Attributes["firstname"].ToString() : "";
                            string lastname = contact.Attributes.ContainsKey("lastname") ? contact.Attributes["lastname"].ToString() : "";
                            string fullname = contact.Attributes.ContainsKey("fullname") ? contact.Attributes["fullname"].ToString() : "";
                            string phone = contact.Attributes.ContainsKey("telephone1") ? contact.Attributes["telephone1"].ToString() : "";
                            string emailaddress = contact.Attributes.ContainsKey("emailaddress1") ? contact.Attributes["emailaddress1"].ToString() : "";
                            string title = contact.Attributes.ContainsKey("jobtitle") ? contact.Attributes["jobtitle"].ToString() : "";
                            string companyname = contact.Attributes.ContainsKey("parentcustomerid") ? ((EntityReference)(contact.Attributes["parentcustomerid"])).Name : "";
                            DateTime modifiedon = contact.Attributes.ContainsKey("modifiedon") ? Convert.ToDateTime(contact.Attributes["modifiedon"]) : DateTime.MinValue;

                            firstname = firstname.Trim(); lastname = lastname.Trim(); fullname = fullname.Trim();

                            string custom = "";

                            bool UpdateContact = true;
                            if (AppSetting.ExchangeUpdateAction.Value.ToUpper() == "DELTA")
                            {
                                if (lastSync > DateTime.MinValue)
                                {
                                    if (lastSync > modifiedon)
                                        UpdateContact = false;
                                }
                            }

                            // Add Code to include Exchange Alias and Revision
                            string alias = contact.Attributes.ContainsKey(AppSetting.ContactExchangeAliasFieldName.Value) ? contact.Attributes[AppSetting.ContactExchangeAliasFieldName.Value].ToString() : "";
                            bool isEmptyAlias = string.IsNullOrEmpty(alias); // ADDED 20170131 to support writing back of missing aliases

                            if (UpdateContact)
                            {
                                if (!string.IsNullOrEmpty(customFieldName))
                                {
                                    if (contact.Attributes.ContainsKey(customFieldName))
                                    {
                                        switch (contact.Attributes[customFieldName].GetType().FullName)
                                        {
                                            case "Microsoft.Xrm.Sdk.EntityReference":
                                                custom = ((EntityReference)(contact.Attributes[customFieldName])).Name;
                                                break;
                                            case "System.String":
                                                custom = contact.Attributes[customFieldName].ToString();
                                                break;
                                            case "System.Int32":
                                                custom = contact.Attributes[customFieldName].ToString();
                                                break;
                                            default:
                                                custom = contact.Attributes[customFieldName].ToString();
                                                break;
                                        }

                                        if (AppSetting.ContactAutoNumberFieldName.Value.Contains(':'))
                                        {
                                            custom = custom.StripPunctuation();

                                            int charCount = 0;
                                            int indexPos = AppSetting.ContactAutoNumberFieldName.Value.IndexOf(':');
                                            bool hasInt = Int32.TryParse(AppSetting.ContactAutoNumberFieldName.Value.Substring(indexPos + 1), out charCount);
                                            if (hasInt)
                                            {
                                                if (charCount > 0)
                                                    custom = (!string.IsNullOrEmpty(custom)) ? custom.Substring(0, charCount) : "";
                                            }
                                        }

                                        if (!string.IsNullOrEmpty(custom))
                                            Trace.AddLog(EventLevel.Verbose, DateTime.Now, "RunSync", "CRM", String.Format("Custom Field Name Value = {0}", custom), "");
                                    }


                                    int stateCode = ((OptionSetValue)(contact.Attributes["statecode"])).Value;

                                    if (stateCode == 0)
                                    {
                                        // bool IsNew = UpdateExchangeContact(emailaddress, ref alias, fullname, firstname, lastname, title, phone, companyname, AppSetting.ExternalDomainName.Value, AppSetting.ExchangeMailContactOU.Value, custom);
                                        bool requireContactUpdate = UpdateExchangeContact(emailaddress, ref alias, fullname, firstname, lastname, title, phone, companyname, custom);
                                        if (requireContactUpdate) // Add code to Update CRM with Alias and Last Sync
                                        {
                                            crm.UpdateContactAlias(contactid, alias); // Optional Field RevisionNumber Removed
                                        }
                                        else
                                        {
                                            if (isEmptyAlias)
                                                crm.UpdateContactAlias(contactid, alias);
                                            else
                                                crm.UpdateContactLastSync(contactid);
                                        }

                                        if (!string.IsNullOrEmpty(alias))
                                        {
                                            if (alias.IsValidEmail())
                                            {
                                                if (AppSetting.DuplicateSyncAction.Value == "MODIFY")
                                                {
                                                    AddDistributionGroupMember(exchangeListName, alias);
                                                }
                                            }
                                            else
                                            {
                                                AddDistributionGroupMember(exchangeListName, alias);
                                            }
                                        }

                                    } // if (stateCode == 0)
                                    else
                                    {
                                        Trace.AddLog(EventLevel.Information, DateTime.Now, "RunSync", "CRM", "ExecuteContactSync", String.Format("Contact {0} Not Added because of Inactive State", fullname), string.Empty, string.Empty);
                                    }
                                } // if (!string.IsNullOrEmpty(customFieldName))
                            }
                            else
                            {
                                Trace.AddLog(EventLevel.Information, DateTime.Now, "RunSync", "CRM", String.Format("Skipping Contact {0} for this type of synchronization.", fullname), string.Empty);
                                AddDistributionGroupMember(exchangeListName, alias);
                            } // if (UpdateContact)

                            if (ProgramSetting.SourceApp == SourceApplication.Windows)
                                ContactSynched(this, new ContactSynchedEventArgs(fullname, emailaddress));

                        } // foreach (Entity contact in contacts.Entities)

                    } // if (AppSetting.FieldMappingAction.Value == "CRM")

                } // if (!stopListProcessing)

            } // if (contacts.Entities.Count > 0)
            else
            {
                exch.RemoveAllDistributionListMembers(exchangeListName);
            } // else if (contacts.Entities.Count > 0)

            crm.UpdateLastExchangeSync(syncListId, DateTime.UtcNow);
        }
      
        public Validator ValidateSync(Guid syncListId, Guid listId, string entityName, string listName, string exchangeGroupName)
        {
            EntityCollection contacts = RetrieveContacts(entityName, listName, listId);
            int totalContact = contacts.Entities.Count;

            List<ExchangeContact> exchangeContacts = exch.GetDistributionListMembers(exchangeGroupName);
            int totalExchangeRecipients = exchangeContacts.Count;

            Validator rc = new Validator(listName, exchangeGroupName, totalContact, totalExchangeRecipients);
            return rc;
        }

        private EntityCollection RetrieveCustomEntityInfo(string entityName, Guid entityId)
        {
            Entity intersect = crm.RetrieveIntersect(entityName);

            string idFieldName = intersect.Contains("xrm_primaryattributename") ? intersect["xrm_primaryattributename"].ToString() : string.Empty; // node.SelectSingleNode("//IdFieldName").InnerText;

            // XmlNode child = node.SelectSingleNode("//IntersectEntity");
            // string intersectType = child.SelectSingleNode("//Type").InnerText;
            string intersectName = intersect.Contains("xrm_intersectentityname") ? intersect["xrm_intersectentityname"].ToString() : string.Empty; // child.SelectSingleNode("//Name").InnerText;
            string intersectParentIdField = intersect.Contains("xrm_parentlookupfieldname") ? intersect["xrm_parentlookupfieldname"].ToString() : string.Empty; // child.SelectSingleNode("//ParentIdField").InnerText;
            Guid intesectParentId = intersect.Contains("xrm_parentlookupfieldid") ? new Guid(intersect["xrm_parentlookupfieldid"].ToString()) : Guid.Empty;
            string intersectContactIdField = intersect.Contains("xrm_contactlookupfieldname") ? intersect["xrm_contactlookupfieldname"].ToString() : string.Empty; // child.SelectSingleNode("//ContactIdField").InnerText;
            bool intersectType = intersect.Contains("xrm_intersecttypecode") ? (bool)intersect["xrm_intersecttypecode"] : false;

            EntityCollection contacts = new EntityCollection();
            if (intersectType == false)
            {
                try
                {
                    contacts = crm.RetrieveManualIntersectMembers(intersectName, intersectContactIdField, intersectParentIdField, entityId);
                }
                catch (System.Exception ex)
                {
                    string message = ex.Message;
                    string command = ex.InnerException.Data.Contains("CommandName") ? ex.InnerException.Data["CommandName"].ToString() : string.Empty;
                    string function = ex.InnerException.Data.Contains("MethodName") ? ex.InnerException.Data["MethodName"].ToString() : string.Empty;

                    List<KeyValuePair<string, string>> parameters = ex.InnerException.Data.Contains("Parameters") ? (List<KeyValuePair<string, string>>)ex.InnerException.Data["Parameters"] : new List<KeyValuePair<string, string>>();
                    string parameterString = Helper.KeyValuePairListToString(parameters);
                    Trace.AddLog(EventLevel.Error, DateTime.Now, "CRM", function, command, message, string.Empty, parameterString);

                }
            }
            else if (intersectType == true)
            {
                // contacts = crm.RetrieveNativeIntersectMembers(entityname, idFieldName, id, intersectName, intersectContactIdField, intersectParentIdField);
            }
            return contacts;
        }

        public bool UpdateExchangeContact(string emailAddress, ref string exchangeAlias, string fullName, string firstName, string lastName, string title, string phoneNumber, string companyName, string customField)
        {
            bool rc = false;

            // Is this an Exchange User
            Trace.AddLog(EventLevel.Verbose, DateTime.Now, "UpdateExchangeContact", "CRM", String.Format("External Domain Value = {0}", AppSetting.ExternalDomainName.Value), "");
            if (emailAddress.ToLower().IsEmailInDomain(AppSetting.ExternalDomainName.Value))
            {
                // Add Company Name Field to Update Exchange User?
                rc = UpdateExchangeUser(emailAddress, ref exchangeAlias, fullName, firstName, lastName, title, phoneNumber);
                return rc;
            }

            // Need Duplicate Alias Section ???

            if (AppSetting.ExchangeIdentityType.Value == "Alias")
            {
                if (!string.IsNullOrEmpty(exchangeAlias))
                {
                    // Alias has value
                    if (ContactExists(exchangeAlias))
                    {
                        Trace.AddLog(EventLevel.Verbose, DateTime.Now, "UpdateExchangeContact", "CRM", "Not a duplicate alias", "");
                        UpdateMailContact(exchangeAlias, fullName, emailAddress);
                        UpdateExchangeContact(exchangeAlias, firstName, lastName, title, phoneNumber, companyName);
                        HideMailContact(exchangeAlias); // Optional String value "HIDE"
                        System.Threading.Thread.Sleep(100);
                    }
                    else if (ContactExists(emailAddress))
                    {
                        Trace.AddLog(EventLevel.Verbose, DateTime.Now, "UpdateExchangeContact", "CRM", "Not a duplicate email address", "");
                        UpdateMailContactAlias(emailAddress, exchangeAlias);
                        UpdateExchangeContact(emailAddress, firstName, lastName, title, phoneNumber, companyName);
                        HideMailContact(emailAddress); // Optional String value "HIDE"
                        System.Threading.Thread.Sleep(100);
                    }
                    else
                    {
                        // Create New Contact
                        CreateNewContact(emailAddress, fullName, firstName, lastName, customField, ref exchangeAlias);
                        UpdateExchangeContact(emailAddress, firstName, lastName, title, phoneNumber, companyName);
                        HideMailContact(emailAddress); // Optional String value "HIDE"
                        System.Threading.Thread.Sleep(100);
                        rc = true;
                    }
                }
                else // Exchange Alias Is Empty
                {
                    // Alias is Empty, Update Contact by Email Address
                    bool isValidEmail = emailAddress.IsValidEmail();
                    if (isValidEmail)
                    {
                        if (ContactExists(emailAddress))
                        {
                            Trace.AddLog(EventLevel.Verbose, DateTime.Now, "UpdateExchangeContact", "CRM", "Not a duplicate email address", "");
                            UpdateMailContact(emailAddress, fullName, emailAddress);
                            UpdateExchangeContact(emailAddress, firstName, lastName, title, phoneNumber, companyName);
                            exchangeAlias = exch.GetMailContact(emailAddress).Alias;
                            HideMailContact(emailAddress); // Optional String value "HIDE"
                            System.Threading.Thread.Sleep(100);
                            rc = true;
                        }
                        else
                        {
                            // Create New Contact
                            CreateNewContact(emailAddress, fullName, firstName, lastName, customField, ref exchangeAlias);
                            UpdateExchangeContact(emailAddress, firstName, lastName, title, phoneNumber, companyName);
                            HideMailContact(emailAddress); // Optional String value "HIDE"
                            System.Threading.Thread.Sleep(100);
                            rc = true;
                        }
                    }
                    else
                    {
                        // Invalid Email address. No Updates Performed
                        Trace.AddLog(EventLevel.Information, DateTime.Now, "UpdateExchangeContact", "CRM", "Invalid Email Address for Sync", String.Format("Email Address {0} is not a valid email address. Please check the contact record for {1}", emailAddress, fullName));
                    }
                }
            }
            else if (AppSetting.ExchangeIdentityType.Value == "EmailAddress")
            {
                bool isValidEmail = emailAddress.IsValidEmail();
                if (isValidEmail)
                {
                    if (ContactExists(emailAddress))
                    {
                        Trace.AddLog(EventLevel.Verbose, DateTime.Now, "UpdateExchangeContact", "CRM", "Not a duplicate email address", "");
                        UpdateMailContact(emailAddress, fullName, emailAddress);
                        UpdateExchangeContact(emailAddress, firstName, lastName, title, phoneNumber, companyName);
                        exchangeAlias = exch.GetMailContact(emailAddress).Alias;
                        HideMailContact(emailAddress); // Optional String value "HIDE"
                        System.Threading.Thread.Sleep(100);
                    }
                    else
                    {
                        // Create New Contact
                        CreateNewContact(emailAddress, fullName, firstName, lastName, customField, ref exchangeAlias);
                        UpdateExchangeContact(emailAddress, firstName, lastName, title, phoneNumber, companyName);
                        HideMailContact(emailAddress); // Optional String value "HIDE"
                        System.Threading.Thread.Sleep(100);
                        rc = true;
                    }
                }
                else
                {
                    // Invalid Email Address - No Updates Performed
                    Trace.AddLog(EventLevel.Information, DateTime.Now, "UpdateExchangeContact", "CRM", "Invalid Email Address for Sync", String.Format("Email Address {0} is not a valid email address. Please check the contact record for {1}", emailAddress, fullName));
                }

            }
            else
            {
                // No Value for Alias or Email Address
                Trace.AddLog(EventLevel.Error, DateTime.Now, "UpdateExchangeContact", "CRM", "Exchange Identity Type does not contain a value");
            }
            return rc;
        }

        public void UpdateExchangeContact(string identity, string firstName, string lastName, string title, string phoneNumber, string companyName)
        {
            List<string> attributes = new List<string>(); List<string> values = new List<string>();
            switch (AppSetting.FieldMappingAction.Value.ToUpper())
            {
                case "NONE":
                    attributes.Add("FirstName"); values.Add(firstName);
                    attributes.Add("LastName"); values.Add(lastName);

                    try
                    {
                        exch.UpdateMailContact(identity, attributes.ToArray(), values.ToArray());
                    }
                    catch (System.Exception ex)
                    {
                        string message = ex.Message;
                        string command = "", function = "", parameterString = "";
                        Helper.LogException(ex, ref command, ref function, ref parameterString);
                        Trace.AddLog(EventLevel.Error, DateTime.Now, "Exchange", function, command, message, string.Empty, parameterString);
                    }

                    break;
                case "MIN":
                    if (!string.IsNullOrEmpty(firstName))
                    {
                        attributes.Add("FirstName"); values.Add(firstName);
                    }

                    if (!string.IsNullOrEmpty(lastName))
                    {
                        attributes.Add("LastName"); values.Add(lastName);
                    }

                    if (!string.IsNullOrEmpty(title))
                    {
                        attributes.Add("Title"); values.Add(title);
                    }

                    if (!string.IsNullOrEmpty(phoneNumber))
                    {
                        attributes.Add("Phone"); values.Add(phoneNumber);
                    }

                    if (!string.IsNullOrEmpty(companyName))
                    {
                        attributes.Add("Company"); values.Add(companyName.MaxLength(64));
                    }

                    try
                    {
                        exch.UpdateMailContact(identity, attributes.ToArray(), values.ToArray());
                    }
                    catch (System.Exception ex)
                    {
                        string message = ex.Message;
                        string command = ex.InnerException.Data["CommandName"].ToString();
                        string function = ex.InnerException.Data["MethodName"].ToString();

                        List<KeyValuePair<string, string>> parameters = (List<KeyValuePair<string, string>>)ex.InnerException.Data["Parameters"];
                        string parameterString = Helper.KeyValuePairListToString(parameters);
                        Trace.AddLog(EventLevel.Error, DateTime.Now, "Exchange", function, command, message, string.Empty, parameterString);

                    }
                    break;
                case "CRM":
                    break;
                default:
                    if (!string.IsNullOrEmpty(firstName))
                    {
                        attributes.Add("FirstName"); values.Add(firstName);
                    }

                    if (!string.IsNullOrEmpty(lastName))
                    {
                        attributes.Add("LastName"); values.Add(lastName);
                    }

                    if (!string.IsNullOrEmpty(title))
                    {
                        attributes.Add("Title"); values.Add(title);
                    }

                    if (!string.IsNullOrEmpty(phoneNumber))
                    {
                        attributes.Add("Phone"); values.Add(phoneNumber);
                    }

                    if (!string.IsNullOrEmpty(companyName))
                    {
                        attributes.Add("Company"); values.Add(companyName.MaxLength(64));
                    }

                    try
                    {
                        exch.UpdateMailContact(identity, attributes.ToArray(), values.ToArray());
                    }
                    catch (System.Exception ex)
                    {
                        string message = ex.Message;
                        string command = ex.InnerException.Data["CommandName"].ToString();
                        string function = ex.InnerException.Data["MethodName"].ToString();

                        List<KeyValuePair<string, string>> parameters = (List<KeyValuePair<string, string>>)ex.InnerException.Data["Parameters"];
                        string parameterString = Helper.KeyValuePairListToString(parameters);
                        Trace.AddLog(EventLevel.Error, DateTime.Now, "Exchange", function, command, message, string.Empty, parameterString);

                    }
                    break;
            }

        }

        public bool UpdateExchangeUser(string emailAddress, ref string exchangeAlias, string fullName, string firstName, string lastName, string title, string phoneNumber)
        {
            bool rc = false;

            if (!string.IsNullOrEmpty(exchangeAlias))
            {
                Trace.AddLog(EventLevel.Information, DateTime.Now, "RunSync", "CRM", String.Format("Add Distribution List Member {0} [{1}]", fullName, exchangeAlias), string.Empty);
                if (ExchangeUserExists(exchangeAlias))
                {
                    Trace.AddLog(EventLevel.Information, DateTime.Now, "RunSync", "Exchange", "Calling Update Mail User Functions", string.Empty);
                    // TODO: NEED TO FIND FIX
                    // The Next Line will update the Mail User Email Address in Exchange
                    // exch.UpdateMailUser(exchangeAlias, fullName, emailAddress);
                    if (!string.IsNullOrEmpty(title))
                    {
                        UpdateExchangeUser(exchangeAlias, new string[] { "FirstName", "LastName", "Phone", "Title" }, new string[] { firstName, lastName, phoneNumber, title });
                    }
                    else
                    {
                        UpdateExchangeUser(exchangeAlias, new string[] { "FirstName", "LastName", "Phone" }, new string[] { firstName, lastName, phoneNumber });
                    }
                    System.Threading.Thread.Sleep(100); 
                }
                else if (ExchangeUserExists(emailAddress))
                {
                    try
                    {
                        ExchangeContact current = exch.GetMailUser(emailAddress);
                        exchangeAlias = current.Alias;
                    }
                    catch (System.Exception ex)
                    {
                        string message = ex.Message;
                        string command = ex.InnerException.Data["CommandName"].ToString();
                        string function = ex.InnerException.Data["MethodName"].ToString();

                        List<KeyValuePair<string, string>> parameters = (List<KeyValuePair<string, string>>)ex.InnerException.Data["Parameters"];
                        string parameterString = Helper.KeyValuePairListToString(parameters);
                        Trace.AddLog(EventLevel.Error, DateTime.Now, "Exchange", function, command, message, string.Empty, parameterString);
                    }

                    Trace.AddLog(EventLevel.Information, DateTime.Now, "RunSync", "Exchange", "Calling Update Mail User Functions", string.Empty);
                    // TODO: NEED TO FIND FIX
                    // The Next Line will update the Mail User Email Address in Exchange
                    // exch.UpdateMailUser(exchangeAlias, fullName, emailAddress);
                    if (!string.IsNullOrEmpty(title))
                    {
                        UpdateExchangeUser(exchangeAlias, new string[] { "FirstName", "LastName", "Phone", "Title" }, new string[] { firstName, lastName, phoneNumber, title });
                    }
                    else
                    {
                        UpdateExchangeUser(exchangeAlias, new string[] { "FirstName", "LastName", "Phone" }, new string[] { firstName, lastName, phoneNumber });
                    }
                    System.Threading.Thread.Sleep(100); 
                    rc = true; // Need to Update Alias in CRM
                }
                else
                {
                    // Will Not Create New Mailbox in Exchange
                    Trace.AddLog(EventLevel.Warning, DateTime.Now, "RunSync", "Exchange", "Mailbox User " + fullName + "does not exist in Exchange", string.Empty);
                    rc = false;
                }
            }
            else
            {
                // alias is empty. If has email address and name, create new contact
                if (!string.IsNullOrEmpty(emailAddress))
                {
                    if (ExchangeUserExists(emailAddress))
                    {
                        try
                        {
                            ExchangeContact current = exch.GetMailUser(emailAddress);
                            exchangeAlias = current.Alias;
                        }
                        catch (System.Exception ex)
                        {
                            string message = ex.Message;
                            string command = ex.InnerException.Data["CommandName"].ToString();
                            string function = ex.InnerException.Data["MethodName"].ToString();

                            List<KeyValuePair<string, string>> parameters = (List<KeyValuePair<string, string>>)ex.InnerException.Data["Parameters"];
                            string parameterString = Helper.KeyValuePairListToString(parameters);
                            Trace.AddLog(EventLevel.Error, DateTime.Now, "Exchange", function, command, message, string.Empty, parameterString);
                        }

                        // TODO: NEED TO FIND FIX
                        // The Next Line will update the Mail User Email Address in Exchange
                        // exch.UpdateMailUser(exchangeAlias, fullName, emailAddress);
                        Trace.AddLog(EventLevel.Information, DateTime.Now, "RunSync", "Exchange", "Calling Update Mail User Functions", string.Empty);
                        if (!string.IsNullOrEmpty(title))
                        {
                            UpdateExchangeUser(exchangeAlias, new string[] { "FirstName", "LastName", "Phone", "Title" }, new string[] { firstName, lastName, phoneNumber, title });
                        }
                        else
                        {
                            UpdateExchangeUser(exchangeAlias, new string[] { "FirstName", "LastName", "Phone" }, new string[] { firstName, lastName, phoneNumber });
                        }
                        System.Threading.Thread.Sleep(100); 
                        rc = true; // Need to Update Alias in CRM
                    }
                    else
                    {
                        // Will Not Create New Mailbox in Exchange
                        Trace.AddLog(EventLevel.Warning, DateTime.Now, "RunSync", "Exchange", "Mailbox User " + fullName + "does not exist in Exchange", string.Empty);
                        rc = false;
                    }
                }
            }
            return rc;
        }

        public bool UpdateExchangeUser(Guid contactid, string emailAddress, ref string exchangeAlias, int revisionNo, string fullName, ref string[] userParams, ref string[] userValues)
        {
            bool rc = false;

            if (!string.IsNullOrEmpty(exchangeAlias))
            {
                Trace.AddLog(EventLevel.Information, DateTime.Now, "RunSync", "CRM", String.Format("Add Distribution List Member {0} [{1}]", fullName, exchangeAlias), string.Empty);
                if (ExchangeUserExists(exchangeAlias))
                {
                    Trace.AddLog(EventLevel.Information, DateTime.Now, "RunSync", "Exchange", "Calling Update Mail User Functions", string.Empty);
                    // TODO: NEED TO FIND FIX
                    // The Next Line will update the Mail User Email Address in Exchange
                    //exch.UpdateMailUser(exchangeAlias, fullName, emailAddress);
                    UpdateExchangeUser(exchangeAlias, userParams, userValues);
                    System.Threading.Thread.Sleep(100); 
                }
                else if (ExchangeUserExists(emailAddress))
                {
                    try
                    {
                        ExchangeContact current = exch.GetMailUser(emailAddress);
                        exchangeAlias = current.Alias;
                    }
                    catch (System.Exception ex)
                    {
                        string message = ex.Message;
                        string command = ex.InnerException.Data["CommandName"].ToString();
                        string function = ex.InnerException.Data["MethodName"].ToString();

                        List<KeyValuePair<string, string>> parameters = (List<KeyValuePair<string, string>>)ex.InnerException.Data["Parameters"];
                        string parameterString = Helper.KeyValuePairListToString(parameters);
                        Trace.AddLog(EventLevel.Error, DateTime.Now, "Exchange", function, command, message, string.Empty, parameterString);
                    }

                    Trace.AddLog(EventLevel.Information, DateTime.Now, "RunSync", "Exchange", "Calling Update Mail User Functions", string.Empty);
                    // TODO: NEED TO FIND FIX
                    // The Next Line will update the Mail User Email Address in Exchange
                    // exch.UpdateMailUser(exchangeAlias, fullName, emailAddress);
                    UpdateExchangeUser(exchangeAlias, userParams, userValues);
                    System.Threading.Thread.Sleep(100); 
                    rc = true; // Need to Update Alias in CRM

                    crm.UpdateContactAlias(contactid, exchangeAlias, revisionNo);
                }
                else
                {
                    // Will Not Create New Mailbox in Exchange
                    Trace.AddLog(EventLevel.Warning, DateTime.Now, "RunSync", "Exchange", "Mailbox User " + fullName + "does not exist in Exchange", string.Empty);
                    rc = false;
                }
            }
            else
            {
                // alias is empty. If has email address and name, create new contact
                if (!string.IsNullOrEmpty(emailAddress))
                {
                    if (ExchangeUserExists(emailAddress))
                    {
                        try
                        {
                            ExchangeContact current = exch.GetMailUser(emailAddress);
                            exchangeAlias = current.Alias;
                        }
                        catch (System.Exception ex)
                        {
                            string message = ex.Message;
                            string command = ex.InnerException.Data["CommandName"].ToString();
                            string function = ex.InnerException.Data["MethodName"].ToString();

                            List<KeyValuePair<string, string>> parameters = (List<KeyValuePair<string, string>>)ex.InnerException.Data["Parameters"];
                            string parameterString = Helper.KeyValuePairListToString(parameters);
                            Trace.AddLog(EventLevel.Error, DateTime.Now, "Exchange", function, command, message, string.Empty, parameterString);
                        }

                        Trace.AddLog(EventLevel.Information, DateTime.Now, "RunSync", "Exchange", "Calling Update Mail User Functions", string.Empty);
                        // TODO: NEED TO FIND FIX
                        // The Next Line will update the Mail User Email Address in Exchange
                        // exch.UpdateMailUser(exchangeAlias, fullName, emailAddress);
                        UpdateExchangeUser(exchangeAlias, userParams, userValues);
                        System.Threading.Thread.Sleep(100); 
                        rc = true; // Need to Update Alias in CRM

                        crm.UpdateContactAlias(contactid, exchangeAlias, revisionNo);

                    }
                    else
                    {
                        // Will Not Create New Mailbox in Exchange
                        Trace.AddLog(EventLevel.Warning, DateTime.Now, "RunSync", "Exchange", "Mailbox User " + fullName + "does not exist in Exchange", string.Empty);
                        rc = false;
                    }
                }
            }
            return rc;

        }

        public bool UpdateExchangeUser(string identity, string[] userParams, string[] userValues)
        {
            if (AppSetting.ExchangeMailUserUpdate.Value.ToUpper() == "YES")
            {
                List<string> attributes = new List<string>(); List<string> values = new List<string>();

                // Add Code to Remove First Name, Last Name and other parameters
                for (int i = 0; i < userParams.Length; i++)
                {
                    // If Attributes Are not First Name or Last Name:
                    switch (userParams[i].ToString())
                    {
                        case "FirstName":
                            if (AppSetting.ExchangeMailUserUpdateFields.Value.Contains(userParams[i].ToString()))
                            {
                                attributes.Add(userParams[i]);
                                values.Add(userValues[i]);
                            }
                            break;
                        case "LastName":
                            if (AppSetting.ExchangeMailUserUpdateFields.Value.Contains(userParams[i].ToString()))
                            {
                                attributes.Add(userParams[i]);
                                values.Add(userValues[i]);
                            }
                            break;
                        case "DisplayName":
                            if (AppSetting.ExchangeMailUserUpdateFields.Value.Contains(userParams[i].ToString()))
                            {
                                attributes.Add(userParams[i]);
                                values.Add(userValues[i]);
                            }
                            break;
                        default:
                            attributes.Add(userParams[i]);
                            values.Add(userValues[i]);
                            break;
                    }
                }

                bool rc = false;
                try
                {
                    rc = exch.UpdateMailUser(identity, attributes.ToArray(), values.ToArray());
                }
                catch (System.Exception ex)
                {
                    string message = ex.Message;
                    string command = ex.InnerException.Data["CommandName"].ToString();
                    string function = ex.InnerException.Data["MethodName"].ToString();

                    List<KeyValuePair<string, string>> parameters = (List<KeyValuePair<string, string>>)ex.InnerException.Data["Parameters"];
                    string parameterString = Helper.KeyValuePairListToString(parameters);
                    Trace.AddLog(EventLevel.Error, DateTime.Now, "Exchange", function, command, message, string.Empty, parameterString);
                }
                return rc;
            }
            else
                return false;

            // return exch.UpdateMailUser(exchangeAlias, userParams, userValues);
        }

        public bool ExchangeUserExists(string identity)
        {
            bool rc = false;
            bool mailboxEnabled = false;
            if (AppSetting.InternalRecipientType.Value.ToUpper() == "MAILBOX-ENABLED")
                mailboxEnabled = true;
            else if (AppSetting.InternalRecipientType.Value.ToUpper() == "MAIL-ENABLED")
                mailboxEnabled = false;
            else
                mailboxEnabled = true;

            try
            {
                rc = exch.ExchangeUserExists(identity, mailboxEnabled);
            }
            catch (System.Exception ex)
            {
                string message = ex.Message;
                string command = ex.InnerException.Data["CommandName"].ToString();
                string function = ex.InnerException.Data["MethodName"].ToString();

                List<KeyValuePair<string, string>> parameters = (List<KeyValuePair<string, string>>)ex.InnerException.Data["Parameters"];
                string parameterString = Helper.KeyValuePairListToString(parameters);
                Trace.AddLog(EventLevel.Error, DateTime.Now, "Exchange", function, command, message, string.Empty, parameterString);
            }

            return rc;
        }

        #endregion

    }
}
