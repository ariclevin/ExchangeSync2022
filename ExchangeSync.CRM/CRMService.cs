using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.ServiceModel;
using System.Text;
using System.Threading.Tasks;

using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Tooling.Connector;

//using Microsoft.Xrm.Client;


namespace ExchangeSync
{
    public class CRMService
    {
        public IOrganizationService service;
        public CrmServiceClient conn;


        public CRMEventArgs Connect(int providerType, NetworkCredential credentials, string crmServer, string organizationName, string regionName)
        {
            switch (providerType)
            {
                case (int)ProviderType.ActiveDirectory:
                    int portNumber = 80;
                    string crmServerName = "";
                    if (crmServer.Contains(':'))
                    {
                        string[] serverName = crmServer.Split(':');
                        crmServerName = serverName[0];
                        int.TryParse(serverName[1], out portNumber);
                    }
                    else
                        crmServerName = crmServer;

                    conn = CRMHelper.Connect(credentials.UserName, credentials.Password, credentials.Domain, crmServerName, portNumber, organizationName, false, false);
                    break;
                case (int)ProviderType.InternetFacingDeployment:
                    conn = CRMHelper.Connect(credentials.UserName, credentials.Password, crmServer, organizationName, false);
                    break;
                case (int)ProviderType.OAuth:
                    break;
                case (int)ProviderType.Office365:
                    conn = CRMHelper.Connect(credentials.UserName, credentials.Password, organizationName, regionName);
                    break;
            }

            if (conn != null)
            {
                if (conn.IsReady)
                {
                    service = (IOrganizationService)conn.OrganizationWebProxyClient != null ? (IOrganizationService)conn.OrganizationWebProxyClient : (IOrganizationService)conn.OrganizationServiceProxy;
                }
            }

            CRMEventArgs args = new CRMEventArgs();
            try
            {
                if (service != null)
                {
                    Guid whoAmI = WhoAmI();
                    if (whoAmI != Guid.Empty)
                    {
                        args.Result = true;
                        args.ConnectionState = CRMEventArgs.CRMState.Connected;
                    }
                    else
                    {
                        args.Result = false;
                        args.ConnectionState = CRMEventArgs.CRMState.NotConnected;
                        args.ConnectionReason = "CRM Connection Error. WhoAmI returned Guid.Empty";
                        args.LastCRMError = conn.LastCrmError;
                    }
                }
                else
                {
                    args.Result = false;
                    args.ConnectionState = CRMEventArgs.CRMState.NotConnected;
                    args.ConnectionReason = "CRM Connection Error. Service was returned as Not Ready";
                    args.LastCRMError = conn.LastCrmError;
                }

            }
            catch (FaultException<OrganizationServiceFault> ex)
            {
                args.Result = false;
                args.ConnectionState = CRMEventArgs.CRMState.NotConnected;
                args.ConnectionReason = "CRM Connection Error: " + ex.Message;
                args.LastCRMError = conn.LastCrmError;
            }
            return args;
        }

        public CRMEventArgs Connect(string crmServerUrl, string clientId, string clientSecret)
        {

            conn = CRMHelper.Connect(crmServerUrl, clientId, clientSecret, true);
            if (conn != null)
            {
                if (conn.IsReady)
                {
                    service = (IOrganizationService)conn.OrganizationWebProxyClient != null ? (IOrganizationService)conn.OrganizationWebProxyClient : (IOrganizationService)conn.OrganizationServiceProxy;
                }
            }

            CRMEventArgs args = new CRMEventArgs();
            try
            {
                if (service != null)
                {
                    Guid whoAmI = WhoAmI();
                    if (whoAmI != Guid.Empty)
                    {
                        args.Result = true;
                        args.ConnectionState = CRMEventArgs.CRMState.Connected;
                    }
                    else
                    {
                        args.Result = false;
                        args.ConnectionState = CRMEventArgs.CRMState.NotConnected;
                        args.ConnectionReason = "CRM Connection Error. WhoAmI returned Guid.Empty";
                        args.LastCRMError = conn.LastCrmError;

                    }
                }
                else
                {
                    args.Result = false;
                    args.ConnectionState = CRMEventArgs.CRMState.NotConnected;
                    args.ConnectionReason = "CRM Connection Error. Service was returned as Not Ready";
                    args.LastCRMError = conn.LastCrmError;

                }

            }
            catch (FaultException<OrganizationServiceFault> ex)
            {
                args.Result = false;
                args.ConnectionState = CRMEventArgs.CRMState.NotConnected;
                args.ConnectionReason = "CRM Connection Error: " + ex.Message;
            }
            return args;
        }


        public CRMEventArgs Connect(string connectionString)
        {

            conn = CRMHelper.Connect(connectionString);
            if (conn != null)
            {
                if (conn.IsReady)
                {
                    service = (IOrganizationService)conn.OrganizationWebProxyClient != null ? (IOrganizationService)conn.OrganizationWebProxyClient : (IOrganizationService)conn.OrganizationServiceProxy;
                }
            }

            CRMEventArgs args = new CRMEventArgs();
            try
            {
                if (service != null)
                {
                    Guid whoAmI = WhoAmI();
                    if (whoAmI != Guid.Empty)
                    {
                        args.Result = true;
                        args.ConnectionState = CRMEventArgs.CRMState.Connected;
                    }
                    else
                    {
                        args.Result = false;
                        args.ConnectionState = CRMEventArgs.CRMState.NotConnected;
                        args.ConnectionReason = "CRM Connection Error. WhoAmI returned Guid.Empty";
                        args.LastCRMError = conn.LastCrmError;

                    }
                }
                else
                {
                    args.Result = false;
                    args.ConnectionState = CRMEventArgs.CRMState.NotConnected;
                    args.ConnectionReason = "CRM Connection Error. Service was returned as Not Ready";
                    args.LastCRMError = conn.LastCrmError;

                }

            }
            catch (FaultException<OrganizationServiceFault> ex)
            {
                args.Result = false;
                args.ConnectionState = CRMEventArgs.CRMState.NotConnected;
                args.ConnectionReason = "CRM Connection Error: " + ex.Message;
            }
            return args;
        }

        public void Disconnect()
        {
            service = null;
        }

        public Guid WhoAmI()
        {
            Guid userid = ((WhoAmIResponse)service.Execute(new WhoAmIRequest())).UserId;
            return userid;

        }

        public List<string> FieldMappings { private get; set; }

        /// <summary>
        /// Replaced RetrieveDistributionLists
        /// </summary>
        /// <param name="SynchOnly"></param>
        /// <returns></returns>
        public EntityCollection RetrieveLists(bool SynchOnly)
        {
            EntityCollection entities = new EntityCollection();
            string[] columnsArray = new string[] { "listid", "listname", "type", "query", "statecode" };

            ColumnSet columns = new ColumnSet(columnsArray);
            QueryExpression query = new QueryExpression
            {
                ColumnSet = columns,
                EntityName = "list",
                Orders =
                {
                    new OrderExpression("listname", OrderType.Ascending)
                }
            };

            try
            {
                entities = service.RetrieveMultiple(query);
            }
            catch (FaultException<OrganizationServiceFault> ex)
            {
                List<KeyValuePair<string, string>> parameters = new List<KeyValuePair<string, string>>()
                {
                    new KeyValuePair<string, string>("xrm_exchangedlname", "NotNull")
                };

                CRMCommandException exc = new CRMCommandException();
                ex.Data.Add("CommandName", "RetrieveMultiple");
                ex.Data.Add("MethodName", "RetrieveLists");
                ex.Data.Add("Parameters", parameters);

                throw new Exception(ex.Message, exc);
            }
            return entities;
        }

        public EntityCollection RetrieveDisconnectedLists()
        {
            EntityCollection entities = new EntityCollection();
            string[] columnsArray = new string[] { "listid", "listname", "type", "xrm_exchangedlname", "xrm_exchangedlstatus", "xrm_exchangelastsync", "query", "statecode" };

            ColumnSet columns = new ColumnSet(columnsArray);
            QueryExpression query = new QueryExpression
            {
                ColumnSet = columns,
                EntityName = "list",
                Orders =
                {
                    new OrderExpression("listname", OrderType.Ascending)
                }
            };

            query.Criteria = new FilterExpression();
            query.Criteria.Conditions.Add(new ConditionExpression("xrm_exchangedlname", ConditionOperator.Null));

            try
            {
                entities = service.RetrieveMultiple(query);
            }
            catch (FaultException<OrganizationServiceFault> ex)
            {
                List<KeyValuePair<string, string>> parameters = new List<KeyValuePair<string, string>>()
                {
                    new KeyValuePair<string, string>("xrm_exchangedlname", "Null")
                };

                CRMCommandException exc = new CRMCommandException();
                ex.Data.Add("CommandName", "RetrieveMultiple");
                ex.Data.Add("MethodName", "RetrieveDisconnectedLists");
                ex.Data.Add("Parameters", parameters);

                throw new Exception(ex.Message, exc);
            }
            return entities;

        }

        // This function retrieves the values for a the SyncList entity
        public EntityCollection RetrieveSyncLists(string profileName, string entityName = "")
        {
            Trace.AddLog(EventLevel.Verbose, DateTime.Now, "CRMService", "RetrieveSyncLists", "Entering Retrieve Sync Lists function");
            EntityCollection entities = new EntityCollection();

            ColumnSet columns = new ColumnSet(true);
            QueryExpression query = new QueryExpression
            {
                ColumnSet = columns,
                EntityName = "xrm_synclist",
                Orders =
                {
                    new OrderExpression("xrm_listname", OrderType.Ascending)
                }
            };

            query.Criteria = new FilterExpression();
            Trace.AddLog(EventLevel.Verbose, DateTime.Now, "CRMService", "RetrieveSyncLists", "Adding Profile Name Condition");
            query.Criteria.Conditions.Add(new ConditionExpression("xrm_profilename", ConditionOperator.Equal, profileName));
            query.Criteria.Conditions.Add(new ConditionExpression("statecode", ConditionOperator.Equal, 0));

            if (!string.IsNullOrEmpty(entityName))
            {
                Trace.AddLog(EventLevel.Verbose, DateTime.Now, "CRMService", "RetrieveSyncLists", "Adding Entity Name Condition");
                query.Criteria.Conditions.Add(new ConditionExpression("xrm_entityname", ConditionOperator.Equal, entityName));
            }

            try
            {
                Trace.AddLog(EventLevel.Verbose, DateTime.Now, "CRMService", "RetrieveSyncLists", "Before Retrieve Multiple");
                entities = service.RetrieveMultiple(query);
                Trace.AddLog(EventLevel.Verbose, DateTime.Now, "CRMService", "RetrieveSyncLists", "After Retrieve Multiple", "Total Entities Retrieved: " + entities.Entities.Count.ToString());
            }
            catch (FaultException<OrganizationServiceFault> ex)
            {
                List<KeyValuePair<string, string>> parameters = new List<KeyValuePair<string, string>>()
                {
                    new KeyValuePair<string, string>("xrm_profilename", profileName),
                    new KeyValuePair<string, string>("xrm_entityname", entityName)
                };

                CRMCommandException exc = new CRMCommandException();
                ex.Data.Add("CommandName", "RetrieveMultiple");
                ex.Data.Add("MethodName", "RetrieveSyncLists");
                ex.Data.Add("Parameters", parameters);


                throw new Exception(ex.Message, exc);
            }
            return entities;
        }

        public EntityCollection RetrieveSyncLists(string profileName, DateTime lastSyncStart)
        {
            Trace.AddLog(EventLevel.Verbose, DateTime.Now, "CRMService", "RetrieveSyncLists", "Entering Retrieve Sync Lists function");
            EntityCollection entities = new EntityCollection();

            ColumnSet columns = new ColumnSet(true);
            QueryExpression query = new QueryExpression
            {
                ColumnSet = columns,
                EntityName = "xrm_synclist",
                Orders =
                {
                    new OrderExpression("xrm_listname", OrderType.Ascending)
                }
            };

            query.Criteria = new FilterExpression();
            Trace.AddLog(EventLevel.Verbose, DateTime.Now, "CRMService", "RetrieveSyncLists", "Adding Profile Name Condition");
            query.Criteria.Conditions.Add(new ConditionExpression("xrm_profilename", ConditionOperator.Equal, profileName));
            query.Criteria.Conditions.Add(new ConditionExpression("xrm_lastsync", ConditionOperator.LessThan, lastSyncStart));
            query.Criteria.Conditions.Add(new ConditionExpression("statecode", ConditionOperator.Equal, 0));

            try
            {
                Trace.AddLog(EventLevel.Verbose, DateTime.Now, "CRMService", "RetrieveSyncLists", "Before Retrieve Multiple");
                entities = service.RetrieveMultiple(query);
                Trace.AddLog(EventLevel.Verbose, DateTime.Now, "CRMService", "RetrieveSyncLists", "After Retrieve Multiple", "Total Entities Retrieved: " + entities.Entities.Count.ToString());
            }
            catch (FaultException<OrganizationServiceFault> ex)
            {
                List<KeyValuePair<string, string>> parameters = new List<KeyValuePair<string, string>>()
                {
                    new KeyValuePair<string, string>("xrm_profilename", profileName),
                    new KeyValuePair<string, string>("xrm_lastsync", lastSyncStart.ToString())
                };

                CRMCommandException exc = new CRMCommandException();
                ex.Data.Add("CommandName", "RetrieveMultiple");
                ex.Data.Add("MethodName", "RetrieveSyncLists");
                ex.Data.Add("Parameters", parameters);


                throw new Exception(ex.Message, exc);
            }
            return entities;
        }


        public EntityCollection RetrieveIntersects()
        {
            EntityCollection entities = new EntityCollection();

            ColumnSet columns = new ColumnSet(true);
            QueryExpression query = new QueryExpression
            {
                ColumnSet = columns,
                EntityName = "xrm_intersect",
                Orders =
                {
                    new OrderExpression("xrm_displayname", OrderType.Ascending)
                }
            };

            query.Criteria = new FilterExpression();
            query.Criteria.Conditions.Add(new ConditionExpression("statecode", ConditionOperator.Equal, 0));


            try
            {
                entities = service.RetrieveMultiple(query);
            }
            catch (FaultException<OrganizationServiceFault> ex)
            {
                List<KeyValuePair<string, string>> parameters = new List<KeyValuePair<string, string>>()
                {
                    new KeyValuePair<string, string>("statecode", "0")
                };

                CRMCommandException exc = new CRMCommandException();
                ex.Data.Add("CommandName", "RetrieveMultiple");
                ex.Data.Add("MethodName", "RetrieveIntersects");
                ex.Data.Add("Parameters", parameters);

                throw new Exception(ex.Message, exc);
            }
            return entities;
        }

        public Entity RetrieveIntersect(Guid intersectId)
        {
            ColumnSet columns = new ColumnSet(true);
            try
            {
                Entity intersect = service.Retrieve("xrm_intersect", intersectId, columns);
                return intersect;
            }
            catch (FaultException<OrganizationServiceFault> ex)
            {
                List<KeyValuePair<string, string>> parameters = new List<KeyValuePair<string, string>>()
                {
                    
                };

                CRMCommandException exc = new CRMCommandException();
                ex.Data.Add("CommandName", "Retrieve");
                ex.Data.Add("MethodName", "RetrieveIntersect");
                ex.Data.Add("Parameters", parameters);

                throw new Exception(ex.Message, exc);
            }
        }

        public Entity RetrieveIntersect(string entityName)
        {
            Entity rc = new Entity();

            ColumnSet columns = new ColumnSet(true);
            QueryExpression query = new QueryExpression
            {
                ColumnSet = columns,
                EntityName = "xrm_intersect",
                Orders =
                {
                    new OrderExpression("xrm_displayname", OrderType.Ascending)
                }
            };

            query.Criteria = new FilterExpression();
            query.Criteria.Conditions.Add(new ConditionExpression("xrm_name", ConditionOperator.Equal, entityName));


            try
            {
                EntityCollection entities = service.RetrieveMultiple(query);
                if (entities.Entities.Count > 0)
                    rc = entities.Entities[0];

            }
            catch (FaultException<OrganizationServiceFault> ex)
            {
                List<KeyValuePair<string, string>> parameters = new List<KeyValuePair<string, string>>()
                {
                    new KeyValuePair<string, string>("statecode", "0")
                };

                CRMCommandException exc = new CRMCommandException();
                ex.Data.Add("CommandName", "RetrieveMultiple");
                ex.Data.Add("MethodName", "RetrieveIntersects");
                ex.Data.Add("Parameters", parameters);

                throw new Exception(ex.Message, exc);
            }
            return rc;

        }

        public Entity RetrieveBusinessEntity(string entityName, Guid entityId)
        {
            return service.Retrieve(entityName, entityId, new ColumnSet(true));
        }

        public EntityCollection RetrieveBusinessEntityCollection(string entityName, string primaryAttributeName, bool SynchOnly)
        {
            EntityCollection entities = new EntityCollection();
            string primaryKeyName = entityName + "id";
            string[] columnsArray = new string[] { primaryKeyName, primaryAttributeName, "statecode" };

            ColumnSet columns = new ColumnSet(columnsArray);
            QueryExpression query = new QueryExpression
            {
                ColumnSet = columns,
                EntityName = entityName,
                Orders =
                {
                    new OrderExpression(primaryAttributeName, OrderType.Ascending)
                }
            };

            if (SynchOnly)
            {
                if (!string.IsNullOrEmpty(ProgramSetting.CRMSolutionNumber))
                {
                    float solutionNumber = 4;
                    bool isNumeric = float.TryParse(ProgramSetting.CRMSolutionNumber, out solutionNumber);
                    if (isNumeric)
                    {
                        if (solutionNumber < 4)
                        {
                            query.Criteria = new FilterExpression();
                            query.Criteria.Conditions.Add(new ConditionExpression("xrm_exchangedlname", ConditionOperator.NotNull));
                        }
                    }
                }
            }

            try
            {
                entities = service.RetrieveMultiple(query);
            }
            catch (FaultException<OrganizationServiceFault> ex)
            {
                List<KeyValuePair<string, string>> parameters = new List<KeyValuePair<string, string>>()
                {
                    new KeyValuePair<string, string>("xrm_exchangedlname", "NotNull")
                };

                CRMCommandException exc = new CRMCommandException();
                ex.Data.Add("CommandName", "RetrieveMultiple");
                ex.Data.Add("MethodName", "RetrieveBusinessEntityCollection");
                ex.Data.Add("Parameters", parameters);

                throw new Exception(ex.Message, exc);
            }
            return entities;
        }

        /// <summary>
        /// Replaced RetrieveDistributionList
        /// </summary>
        /// <param name="SynchOnly"></param>
        /// <returns></returns>
        public Entity RetrieveList(Guid listid)
        {
            string[] columnsArray = new string[] { "listid", "listname", "type", "query", "statecode" };
            ColumnSet columns = new ColumnSet(columnsArray);
            return service.Retrieve("list", listid, columns);
        }

        public void RemoveFromSyncList(Guid id)
        {
            try
            {
                service.Delete("xrm_synclist", id);
            }
            catch (FaultException<OrganizationServiceFault> ex)
            {
                List<KeyValuePair<string, string>> parameters = new List<KeyValuePair<string, string>>()
                {
                    new KeyValuePair<string, string>("xrm_synclistid", id.ToString())
                };

                CRMCommandException exc = new CRMCommandException();
                ex.Data.Add("CommandName", "Delete");
                ex.Data.Add("MethodName", "RemoveFromSyncList");
                ex.Data.Add("Parameters", parameters);

                throw new Exception(ex.Message, exc);
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="id"></param>
        /// <param name="entityname"></param>
        /// <param name="status">0 = Inactive, 1 = Active</param>
        public void ChangeSyncListStatus(Guid id, string entityname, bool status)
        {
            Entity entity = new Entity(entityname);
            entity.Id = id;
            entity.Attributes["xrm_exchangedlstatus"] = status;
            try
            {
                service.Update(entity);
            }
            catch (FaultException<OrganizationServiceFault> ex)
            {
                List<KeyValuePair<string, string>> parameters = new List<KeyValuePair<string, string>>()
                {
                    new KeyValuePair<string, string>("xrm_exchangedlstatus", ""),
                };

                CRMCommandException exc = new CRMCommandException();
                ex.Data.Add("CommandName", "Update");
                ex.Data.Add("MethodName", "ChangeSyncListStatus");
                ex.Data.Add("Parameters", parameters);

                throw new Exception(ex.Message, exc);
            }
        }

        public void ChangeSyncListStatus(Guid id, bool status)
        {
            Entity entity = new Entity("xrm_synclist");
            entity.Id = id;
            entity.Attributes["xrm_liststatus"] = status;
            entity.Attributes["xrm_autosyncstatus"] = status;

            try
            {
                service.Update(entity);
            }
            catch (FaultException<OrganizationServiceFault> ex)
            {
                List<KeyValuePair<string, string>> parameters = new List<KeyValuePair<string, string>>()
                {
                    new KeyValuePair<string, string>("xrm_liststatus", status.ToString()),
                    new KeyValuePair<string, string>("xrm_autosyncstatus", status.ToString())
                };

                CRMCommandException exc = new CRMCommandException();
                ex.Data.Add("CommandName", "Update");
                ex.Data.Add("MethodName", "ChangeSyncListStatus");
                ex.Data.Add("Parameters", parameters);

                throw new Exception(ex.Message, exc);
            }
        }

        public void AddToSyncList(Guid id, string entityname, string dlname)
        {
            Entity entity = new Entity(entityname);
            entity.Id = id;
            entity.Attributes["xrm_exchangedlname"] = dlname;
            entity.Attributes["xrm_exchangedlstatus"] = true;
            try
            {
                service.Update(entity);
            }
            catch (FaultException<OrganizationServiceFault> ex)
            {
                List<KeyValuePair<string, string>> parameters = new List<KeyValuePair<string, string>>()
                {
                    new KeyValuePair<string, string>("xrm_exchangedlname", dlname),
                    new KeyValuePair<string, string>("xrm_exchangedlstatus", "true")
                };

                CRMCommandException exc = new CRMCommandException();
                ex.Data.Add("CommandName", "Update");
                ex.Data.Add("MethodName", "AddToSyncList");
                ex.Data.Add("Parameters", parameters);

                throw new Exception(ex.Message, exc);
            }
        }

        public void AddToSyncList(Guid id, string entityname, string dlname, string ou)
        {
            Entity entity = new Entity(entityname);
            entity.Id = id;
            entity.Attributes["xrm_exchangedlname"] = dlname;
            entity.Attributes["xrm_exchangeoucanonicalname"] = ou;
            entity.Attributes["xrm_exchangedlstatus"] = true;
            try
            {
                service.Update(entity);
            }
            catch (FaultException<OrganizationServiceFault> ex)
            {
                List<KeyValuePair<string, string>> parameters = new List<KeyValuePair<string, string>>()
                {
                    new KeyValuePair<string, string>("xrm_exchangedlname", dlname),
                    new KeyValuePair<string, string>("xrm_exchangeoucanonicalname", ou),
                    new KeyValuePair<string, string>("xrm_exchangedlstatus", "true")
                };

                CRMCommandException exc = new CRMCommandException();
                ex.Data.Add("CommandName", "Update");
                ex.Data.Add("MethodName", "AddToSyncList");
                ex.Data.Add("Parameters", parameters);

                throw new Exception(ex.Message, exc);
            }
        }

        public Entity RetrieveContact(Guid contactId, string customFieldName = "", string revisionFieldName = "")
        {
            ColumnSet columns = new ColumnSet(new string[] { "contactid", "firstname", "lastname", "fullname", "emailaddress1", "xrm_exchangealias", "modifiedon" });

            if (!string.IsNullOrEmpty(AppSetting.ContactAutoNumberFieldName.Value))
                columns.AddColumn(AppSetting.ContactAutoNumberFieldName.Value);

            if (!string.IsNullOrEmpty(AppSetting.ContactRevisionFieldName.Value))
                columns.AddColumn(AppSetting.ContactRevisionFieldName.Value);

            Entity contact = service.Retrieve("contact", contactId, columns);
            return contact;
        }

        public EntityCollection RetrieveContactsWithAlias()
        {
            EntityCollection rc = new EntityCollection();
            QueryExpression query = new QueryExpression()
            {
                EntityName = "contact",
                ColumnSet = new ColumnSet(new string[] { "contactid", "fullname", "xrm_exchangealias", "xrm_revision", "modifiedon" }),
                Criteria =
                {
                    Conditions =
                    {
                        new ConditionExpression("xrm_exchangealias", ConditionOperator.NotNull)
                    }
                }
            };

            try
            {
                rc = service.RetrieveMultiple(query);
            }
            catch (FaultException<OrganizationServiceFault> ex)
            {
                List<KeyValuePair<string, string>> parameters = new List<KeyValuePair<string, string>>()
                {
                    new KeyValuePair<string, string>("xrm_exchangealias", "NotNull")
                };

                CRMCommandException exc = new CRMCommandException();
                ex.Data.Add("CommandName", "RetrieveMultiple");
                ex.Data.Add("MethodName", "RetrieveContactsWithAlias");
                ex.Data.Add("Parameters", parameters);

                throw new Exception(ex.Message, exc);
            }

            return rc;
        }

        public EntityCollection RetrieveContactListMembers(Guid listId)
        {
            ColumnSet columns = new ColumnSet("contactid", "firstname", "lastname", "fullname", "telephone1", "jobtitle", "emailaddress1", "parentcustomerid", "statecode", "modifiedon");
            return RetrieveListMembers(columns, listId, "contact", "contactid");
        }

        public EntityCollection RetrieveLeadListMembers(Guid listId)
        {
            ColumnSet columns = new ColumnSet("leadid", "firstname", "lastname", "fullname", "telephone1", "jobtitle", "emailaddress1", "parentcustomerid", "statecode", "modifiedon");
            return RetrieveListMembers(columns, listId, "lead", "leadid");
        }

        public EntityCollection RetrieveAccountListMembers(Guid listId)
        {
            ColumnSet columns = new ColumnSet("accountid", "name", "telephone1", "emailaddress1", "statecode", "modifiedon");
            return RetrieveListMembers(columns, listId, "lead", "leadid");
        }


        private EntityCollection RetrieveListMembers(ColumnSet columns, Guid listId, string entityName, string linkFromAttributeName)
        {
            if ((FieldMappings != null) && (FieldMappings.Count > 0))
            {
                foreach (string field in FieldMappings)
                {
                    if (!columns.Columns.Contains(field))
                        columns.AddColumn(field);
                }
            }

            return RetrieveAllListMembers(columns, listId, entityName, linkFromAttributeName);
        }

        public EntityCollection RetrieveAllListMembers(ColumnSet columns, Guid listId, string entityName, string linkFromAttributeName)
        {
            EntityCollection rc = new EntityCollection();
            QueryExpression query = new QueryExpression()
            {
                EntityName = "contact",
                ColumnSet = columns,
                LinkEntities =
                {
                   new LinkEntity()
                   {
                       LinkToEntityName = "listmember",
                       LinkToAttributeName = "entityid",
                       LinkFromAttributeName = "contactid",
                       LinkEntities =
                       {
                           new LinkEntity()
                           {
                               LinkToEntityName = "list",
                               LinkToAttributeName = "listid",
                               LinkFromAttributeName = "listid",
                               LinkCriteria =
                               {
                                    Conditions =
                                    {
                                        new ConditionExpression("listid", ConditionOperator.Equal, listId)
                                    }
                               }
                           }
                       }
                   }
                }
            };

            if (!string.IsNullOrEmpty(AppSetting.ContactRequiredAttributes.Value))
            {
                if (AppSetting.ContactRequiredAttributes.Value.Contains(";"))
                {
                    string[] customColumns = AppSetting.ContactRequiredAttributes.Value.Split(new char[] { ';' });
                    foreach (string customColumn in customColumns)
                    {
                        if (!query.ColumnSet.Columns.Contains(customColumn))
                            query.ColumnSet.AddColumns(customColumn);
                    }
                }
                else
                {
                    query.ColumnSet.AddColumn(AppSetting.ContactRequiredAttributes.Value);
                }
            }

            if ((FieldMappings != null) && (FieldMappings.Count > 0))
            {
                foreach (string field in FieldMappings)
                {
                    if (!query.ColumnSet.Columns.Contains(field))
                        query.ColumnSet.AddColumn(field);
                }
            }

            try
            {
                rc = service.RetrieveMultiple(query);
            }
            catch (FaultException<OrganizationServiceFault> ex)
            {
                List<KeyValuePair<string, string>> parameters = new List<KeyValuePair<string, string>>()
                {
                    new KeyValuePair<string, string>("listid", listId.ToString())
                };

                CRMCommandException exc = new CRMCommandException();
                ex.Data.Add("CommandName", "RetrieveMultiple");
                ex.Data.Add("MethodName", "RetrieveListMembers");
                ex.Data.Add("Parameters", parameters);

                throw new Exception(ex.Message, exc);

                // Logger.AddLog(EventLevel.Information, DateTime.Now, "RetrieveListMembers", "CRM", errColumnList, "");
                // Logger.AddLog(EventLevel.Error, DateTime.Now, "RetrieveListMembers", "CRM", ex.Message, "");
            }
            return rc;
        }

        public EntityCollection RetrieveListMembers(string query, bool convertToQueryExpression)
        {
            EntityCollection rc = new EntityCollection();

            if (convertToQueryExpression)
            {
                rc = RetrieveListMembers(query);
            }
            else
            {
                FetchExpression fe = new FetchExpression(query);
                try
                {
                    rc = service.RetrieveMultiple(fe);
                }
                catch (FaultException<OrganizationServiceFault> ex)
                {
                    List<KeyValuePair<string, string>> parameters = new List<KeyValuePair<string, string>>()
                    {
                        new KeyValuePair<string, string>("query", query)
                    };

                    CRMCommandException exc = new CRMCommandException();
                    ex.Data.Add("CommandName", "RetrieveMultiple");
                    ex.Data.Add("MethodName", "RetrieveListMembers");
                    ex.Data.Add("Parameters", parameters);

                    throw new Exception(ex.Message, exc);
                }
            }
            return rc;
        }

        private EntityCollection RetrieveListMembers(string fetchxml)
        {
            EntityCollection rc = new EntityCollection();
            FetchXmlToQueryExpressionRequest request = new FetchXmlToQueryExpressionRequest { FetchXml = fetchxml };
            FetchXmlToQueryExpressionResponse response = (FetchXmlToQueryExpressionResponse)service.Execute(request);

            QueryExpression query = response.Query;

            if (!string.IsNullOrEmpty(AppSetting.ContactRequiredAttributes.Value))
            {
                if (AppSetting.ContactRequiredAttributes.Value.Contains(";"))
                {
                    string[] customColumns = AppSetting.ContactRequiredAttributes.Value.Split(new char[] { ';' });
                    query.ColumnSet.AddColumns(customColumns);
                }
                else
                {
                    query.ColumnSet.AddColumn(AppSetting.ContactRequiredAttributes.Value);
                }
            }

            if ((FieldMappings != null) && (FieldMappings.Count > 0))
            {
                foreach (string field in FieldMappings)
                {
                    if (!query.ColumnSet.Columns.Contains(field))
                        query.ColumnSet.AddColumn(field);
                }
            }

            // Add Code to Load All of the Columns that are required by the application 8/19/2015
            string[] requiredColumns = new string[] { "contactid", "firstname", "lastname", "fullname", "telephone1", "jobtitle", "emailaddress1", "parentcustomerid", "statecode", "modifiedon" };
            foreach (string field in requiredColumns)
            {
                if (!query.ColumnSet.Columns.Contains(field))
                    query.ColumnSet.AddColumn(field);
            }

            try
            {
                rc = service.RetrieveMultiple(query);
            }
            catch (FaultException<OrganizationServiceFault> ex)
            {
                List<KeyValuePair<string, string>> parameters = new List<KeyValuePair<string, string>>()
                {
                    new KeyValuePair<string, string>("fetchXml", fetchxml)
                };

                CRMCommandException exc = new CRMCommandException();
                ex.Data.Add("CommandName", "RetrieveMultiple");
                ex.Data.Add("MethodName", "RetrieveListMembers");
                ex.Data.Add("Parameters", parameters);

                throw new Exception(ex.Message, exc);
            }
            return rc;

        }

        public EntityCollection RetrieveManualIntersectMembers(string linkToEntityName, string linkToAttributeName, string parentEntityAttributeName, Guid parentEntityAttributeId)
        {
            ColumnSet columns = new ColumnSet(new string[] { "contactid", "firstname", "lastname", "fullname", "telephone1", "jobtitle", "emailaddress1", "parentcustomerid", "statecode", "modifiedon" });
            if ((FieldMappings != null) && (FieldMappings.Count > 0))
            {
                foreach (string field in FieldMappings)
                {
                    if (!columns.Columns.Contains(field))
                        columns.AddColumn(field);
                }
            }
            return RetrieveManualIntersectMembers(linkToEntityName, linkToAttributeName, parentEntityAttributeName, parentEntityAttributeId, columns);
        }
       
        private EntityCollection RetrieveManualIntersectMembers(string linkToEntityName, string linkToAttributeName, string parentEntityAttributeName, Guid parentEntityAttributeId, ColumnSet columns)
        {
            EntityCollection rc = new EntityCollection();
            QueryExpression query = new QueryExpression()
            {
                EntityName = "contact",
                ColumnSet = columns,
                LinkEntities =
                {
                   new LinkEntity()
                   {
                       LinkToEntityName = linkToEntityName,
                       LinkToAttributeName = linkToAttributeName,
                       LinkFromEntityName = "contact",
                       LinkFromAttributeName = "contactid",
                       LinkCriteria =
                       {
                            Conditions =
                            {
                                new ConditionExpression(parentEntityAttributeName, ConditionOperator.Equal, parentEntityAttributeId),
                                new ConditionExpression("statecode", ConditionOperator.Equal, 0)
                            }
                       },
                   }
                }
            };

            if (!string.IsNullOrEmpty(AppSetting.ContactRequiredAttributes.Value))
            {
                if (AppSetting.ContactRequiredAttributes.Value.Contains(";"))
                {
                    string[] customColumns = AppSetting.ContactRequiredAttributes.Value.Split(new char[] { ';' });
                    query.ColumnSet.AddColumns(customColumns);
                }
                else
                {
                    query.ColumnSet.AddColumn(AppSetting.ContactRequiredAttributes.Value);
                }
            }

            if ((FieldMappings != null) && (FieldMappings.Count > 0))
            {
                foreach (string field in FieldMappings)
                {
                    if (!query.ColumnSet.Columns.Contains(field))
                        query.ColumnSet.AddColumn(field);
                }
            }

            try
            {
                rc = service.RetrieveMultiple(query);
            }
            catch (FaultException<OrganizationServiceFault> ex)
            {
                List<KeyValuePair<string, string>> parameters = new List<KeyValuePair<string, string>>()
                {
                    new KeyValuePair<string, string>(parentEntityAttributeName, parentEntityAttributeId.ToString()),
                    new KeyValuePair<string, string>("statecode", "0")
                };

                CRMCommandException exc = new CRMCommandException();
                ex.Data.Add("CommandName", "RetrieveMultiple");
                ex.Data.Add("MethodName", "RetrieveManualIntersectMembers");
                ex.Data.Add("Parameters", parameters);

                throw new Exception(ex.Message, exc);
            }
            return rc;
        }

        public EntityCollection RetrieveNativeIntersectMembers(string parentEntityName, string parentEntityIdAttribute, Guid parentEntityId, string intersectEntityName, string intersectEntityContactIdAttribute, string intersectEntityRelatedIdAttribute)
        {
            ColumnSet columns = new ColumnSet(new string[] { "contactid", "firstname", "lastname", "fullname", "telephone1", "jobtitle", "emailaddress1", "parentcustomerid", "statecode", "modifiedon" });
            foreach (string field in FieldMappings)
            {
                if (!columns.Columns.Contains(field))
                    columns.AddColumn(field);
            }

            return RetrieveNativeIntersectMembers(parentEntityName, parentEntityIdAttribute, parentEntityId, intersectEntityName, intersectEntityContactIdAttribute, intersectEntityRelatedIdAttribute, columns);
        }

        private EntityCollection RetrieveNativeIntersectMembers(string parentEntityName, string parentEntityIdAttribute, Guid parentEntityId, string intersectEntityName, string intersectEntityContactIdAttribute, string intersectEntityRelatedIdAttribute, ColumnSet columns)
        {
            EntityCollection rc = new EntityCollection();
            QueryExpression query = new QueryExpression()
            {
                EntityName = "contact",
                ColumnSet = columns,
                LinkEntities =
                {
                   new LinkEntity() // Intersect Entity
                   {
                       LinkToEntityName = intersectEntityName, // xrm_entityName
                       LinkToAttributeName = intersectEntityContactIdAttribute, // xrm_contactid
                       LinkFromAttributeName = "contactid", // OK
                       LinkEntities =
                       {
                           new LinkEntity() // Related Entity
                           {
                               LinkToEntityName = parentEntityName,
                               LinkToAttributeName = parentEntityIdAttribute,
                               LinkFromAttributeName = intersectEntityRelatedIdAttribute, // Intersect Entity xrm_listid
                               LinkCriteria =
                               {
                                    Conditions =
                                    {
                                        new ConditionExpression(parentEntityIdAttribute, ConditionOperator.Equal, parentEntityId)
                                    }
                               }
                           }
                       }
                   }
                }
            };

            if (!string.IsNullOrEmpty(AppSetting.ContactRequiredAttributes.Value))
            {
                if (AppSetting.ContactRequiredAttributes.Value.Contains(";"))
                {
                    string[] customColumns = AppSetting.ContactRequiredAttributes.Value.Split(new char[] { ';' });
                    query.ColumnSet.AddColumns(customColumns);
                }
                else
                {
                    query.ColumnSet.AddColumn(AppSetting.ContactRequiredAttributes.Value);
                }
            }

            if ((FieldMappings != null) && (FieldMappings.Count > 0))
            {
                foreach (string field in FieldMappings)
                {
                    if (!columns.Columns.Contains(field))
                        columns.AddColumn(field);
                }
            }

            try
            {
                rc = service.RetrieveMultiple(query);
            }
            catch (FaultException<OrganizationServiceFault> ex)
            {
                List<KeyValuePair<string, string>> parameters = new List<KeyValuePair<string, string>>()
                {
                    new KeyValuePair<string, string>(parentEntityIdAttribute, parentEntityId.ToString())
                };

                CRMCommandException exc = new CRMCommandException();
                ex.Data.Add("CommandName", "RetrieveMultiple");
                ex.Data.Add("MethodName", "RetrieveNativeIntersectMembers");
                ex.Data.Add("Parameters", parameters);

                throw new Exception(ex.Message, exc);
            }
            return rc;
        }

        public EntityCollection RetrieveChangeTrackingOperations(int objectTypeCode)
        {
            ColumnSet columns = new ColumnSet("xrm_changetrackingoperationid", "xrm_alias", "xrm_emailaddress");

            QueryExpression query = new QueryExpression()
            {
                EntityName = "xrm_changetrackingoperation",
                ColumnSet = columns,
                Criteria =
                {
                    Conditions =
                    {
                        new ConditionExpression("xrm_syncobjecttypecode", ConditionOperator.Equal, objectTypeCode),
                        new ConditionExpression("statecode", ConditionOperator.Equal, 0)
                    }
                }
            };

            EntityCollection results = service.RetrieveMultiple(query);
            return results;
        }

        private void UpdateAlias(Guid entityId, string entityName, string alias, int revision = 0)
        {
            Entity entity = new Entity(entityName);
            entity.Id = entityId;
            entity.Attributes["xrm_exchangealias"] = alias;
            entity.Attributes["xrm_lastsync"] = DateTime.UtcNow;

            if (revision >= 0)
            {
                revision++;
                entity.Attributes["xrm_revision"] = revision;
            }

            try
            {
                service.Update(entity);
            }
            catch (FaultException<OrganizationServiceFault> ex)
            {
                List<KeyValuePair<string, string>> parameters = new List<KeyValuePair<string, string>>()
                {
                    new KeyValuePair<string, string>("xrm_exchangealias", alias),
                    new KeyValuePair<string, string>("xrm_lastsync", DateTime.UtcNow.ToString())
                };

                CRMCommandException exc = new CRMCommandException();
                ex.Data.Add("CommandName", "Update");
                ex.Data.Add("MethodName", "UpdateContactAlias");
                ex.Data.Add("Parameters", parameters);

                throw new Exception(ex.Message, exc);
                // Logger.AddLog(EventLevel.Error, DateTime.Now, "UpdateContactAlias", "CRM", String.Format("Error Calling UpdateContactAlias for alias {0} ", alias), ex.Message);
            }

        }

        public void UpdateContactAlias(Guid contactId, string alias, int revision = 0)
        {
            UpdateAlias(contactId, "contact", alias, revision);
        }

        public void UpdateLeadAlias(Guid leadId, string alias, int revision = 0)
        {
            UpdateAlias(leadId, "lead", alias, revision);
        }

        public void UpdateAccountAlias(Guid accountId, string alias, int revision = 0)
        {
            UpdateAlias(accountId, "account", alias, revision);
        }

        private void UpdateLastSync(Guid entityId, string entityName)
        {
            Entity entity = new Entity(entityName);
            entity.Id = entityId;
            entity.Attributes["xrm_lastsync"] = DateTime.UtcNow;

            try
            {
                service.Update(entity);
            }
            catch (FaultException<OrganizationServiceFault> ex)
            {
                List<KeyValuePair<string, string>> parameters = new List<KeyValuePair<string, string>>()
                {
                    new KeyValuePair<string, string>("xrm_lastsync", DateTime.UtcNow.ToString()),
                    new KeyValuePair<string, string>("entityName", entityName),
                };

                CRMCommandException exc = new CRMCommandException();
                ex.Data.Add("CommandName", "Update");
                ex.Data.Add("MethodName", "UpdateLastSync");
                ex.Data.Add("Parameters", parameters);

                throw new Exception(ex.Message, exc);
            }
        }


        public void UpdateContactLastSync(Guid contactId)
        {
            UpdateLastSync(contactId, "contact");
        }

        public void UpdateLeadLastSync(Guid leadId)
        {
            UpdateLastSync(leadId, "lead");
        }

        public void UpdateAccountLastSync(Guid accountId)
        {
            UpdateLastSync(accountId, "account");
        }


        public void UpdateChangeTrackingOperation(Guid changeTrackingOperationId, string alias, string emailAddress)
        {
            Entity cto = new Entity("xrm_changetrackingoperation");
            cto.Id = changeTrackingOperationId;
            cto.Attributes["statecode"] = new OptionSetValue(1);
            cto.Attributes["statuscode"] = new OptionSetValue(2);

            try
            {
                service.Update(cto);
            }
            catch (FaultException<OrganizationServiceFault> ex)
            {
                List<KeyValuePair<string, string>> parameters = new List<KeyValuePair<string, string>>()
                {
                    new KeyValuePair<string, string>("id", changeTrackingOperationId.ToString()),
                    new KeyValuePair<string, string>("alias", alias),
                    new KeyValuePair<string, string>("email", emailAddress)
                };


                CRMCommandException exc = new CRMCommandException();
                ex.Data.Add("CommandName", "Update");
                ex.Data.Add("MethodName", "UpdateChangeTrackingOperation");
                ex.Data.Add("Parameters", parameters);

                throw new Exception(ex.Message, exc);
            }
        }

        public void DeleteChangeTrackOperation(Guid changeTrackingOperationId, string alias, string emailAddress)
        {
            try
            {
                service.Delete("xrm_changetrackingoperation", changeTrackingOperationId);
            }
            catch (FaultException<OrganizationServiceFault> ex)
            {
                List<KeyValuePair<string, string>> parameters = new List<KeyValuePair<string, string>>()
                {
                    new KeyValuePair<string, string>("id", changeTrackingOperationId.ToString()),
                    new KeyValuePair<string, string>("alias", alias),
                    new KeyValuePair<string, string>("email", emailAddress)
                };


                CRMCommandException exc = new CRMCommandException();
                ex.Data.Add("CommandName", "Delete");
                ex.Data.Add("MethodName", "DeleteChangeTrackOperation");
                ex.Data.Add("Parameters", parameters);

                throw new Exception(ex.Message, exc);
            }
        }


        #region Administration Functions

        public void UpdateLastExchangeSync(Guid syncListId, DateTime lastExchangeSync)
        {
            Entity syncList = new Entity("xrm_synclist");
            syncList.Id = syncListId;
            syncList.Attributes["xrm_lastsync"] = lastExchangeSync;

            try
            {
                service.Update(syncList);
            }
            catch (FaultException<OrganizationServiceFault> ex)
            {
                List<KeyValuePair<string, string>> parameters = new List<KeyValuePair<string, string>>()
                {
                    new KeyValuePair<string, string>("xrm_lastsync", lastExchangeSync.ToString())
                };

                CRMCommandException exc = new CRMCommandException();
                ex.Data.Add("CommandName", "Update");
                ex.Data.Add("MethodName", "UpdateLastExchangeSync");
                ex.Data.Add("Parameters", parameters);

                throw new Exception(ex.Message, exc);
                // Trace.AddLog(EventLevel.Error, DateTime.Now, "UpdateLastExchangeSync", "CRM", "Error Calling UpdateLastExchangeSync for SyncList ", ex.Message);
            }
        }

        public void UpdateLastExchangeSync(string entityName, Guid entityId, DateTime lastExchangeSync)
        {
            Entity custom = new Entity(entityName);
            custom.Id = entityId;
            custom.Attributes["xrm_exchangelastsync"] = lastExchangeSync;

            try
            {
                service.Update(custom);
            }
            catch (FaultException<OrganizationServiceFault> ex)
            {
                List<KeyValuePair<string, string>> parameters = new List<KeyValuePair<string, string>>()
                {
                    new KeyValuePair<string, string>("xrm_lastsync", lastExchangeSync.ToString())
                };

                CRMCommandException exc = new CRMCommandException();
                ex.Data.Add("CommandName", "Update");
                ex.Data.Add("MethodName", "UpdateLastExchangeSync");
                ex.Data.Add("Parameters", parameters);

                throw new Exception(ex.Message, exc);
                // Logger.AddLog(EventLevel.Error, DateTime.Now, "UpdateLastExchangeSync", "CRM", String.Format("Error Calling UpdateLastExchangeSync for entity {0} ", entityName), ex.Message);
            }
        }

        private BulkDetectDuplicatesResponse _response; private Entity _rule; //DuplicateRule
        public List<Guid> CheckDuplicateEmails(bool useExistingRule = false, float minuteCalculation = 1)
        {
            int numSeconds = Convert.ToInt32(minuteCalculation * 1 * 60);
            if (!useExistingRule)
                CreateDuplicateDetectionRule(numSeconds);

            BulkDetectDuplicatesRequest request = new BulkDetectDuplicatesRequest()
            {
                JobName = "Detect Duplicate Contacts",
                Query = new QueryExpression()
                {
                    EntityName = "contact",
                    ColumnSet = new ColumnSet(true)
                },
                RecurrencePattern = String.Empty,
                RecurrenceStartTime = DateTime.Now,
                ToRecipients = new Guid[0],
                CCRecipients = new Guid[0]
            };

            _response = (BulkDetectDuplicatesResponse)service.Execute(request);

            numSeconds = Convert.ToInt32(minuteCalculation * 2 * 60);
            WaitForAsyncJobToFinish(_response.JobId, numSeconds);

            QueryByAttribute query = new QueryByAttribute()
            {
                ColumnSet = new ColumnSet(true),
                EntityName = "duplicaterecord"
            };
            query.Attributes.Add("asyncoperationid");
            query.Values.Add(_response.JobId);
            EntityCollection results = service.RetrieveMultiple(query);

            var duplicateIds = results.Entities.Select((entity) =>
                ((EntityReference)(entity.Attributes["baserecordid"])).Id);

            if (!useExistingRule)
                service.Delete("duplicaterule", _rule.Id);

            return duplicateIds.ToList<Guid>();
        }

        private void WaitForAsyncJobToFinish(Guid jobId, int maxTimeSeconds)
        {
            for (int i = 0; i < maxTimeSeconds; i++)
            {
                Entity asyncJob = service.Retrieve("asyncoperation",
                    jobId, new ColumnSet("statecode"));

                if ((int)asyncJob.Attributes["statecode"] == ExchangeSync.AsyncOperationState.Completed.ToInt())
                    return;

                System.Threading.Thread.Sleep(1000);
            }
            throw new Exception(String.Format(
                "  Exceeded maximum time of {0} seconds waiting for asynchronous job to complete",
                maxTimeSeconds
            ));
        }

        private void CreateDuplicateDetectionRule(int numSeconds)
        {
            _rule = new Entity("duplicaterule");
            _rule.Attributes["name"] = "Contacts with the same Email Address";
            _rule.Attributes["baseentityname"] = "contact";
            _rule.Attributes["matchingentityname"] = "contact";
            _rule.Id = service.Create(_rule);

            Entity emailCondition = new Entity("duplicaterulecondition");
            emailCondition.Attributes.Add(new KeyValuePair<string, object>("baseattributename", "emailaddress1"));
            emailCondition.Attributes.Add(new KeyValuePair<string, object>("matchingattributename", "emailaddress1"));
            emailCondition.Attributes.Add(new KeyValuePair<string, object>("operatorcode", new OptionSetValue(0)));
            emailCondition.Attributes.Add(new KeyValuePair<string, object>("regardingobjectid", _rule.ToEntityReference()));
            service.Create(emailCondition);

            PublishDuplicateRuleRequest publishRequest = new PublishDuplicateRuleRequest()
            {
                DuplicateRuleId = _rule.Id
            };
            PublishDuplicateRuleResponse publishResponse = (PublishDuplicateRuleResponse)service.Execute(publishRequest);
            WaitForAsyncJobToFinish(publishResponse.JobId, numSeconds);
        }

        public EntityCollection RetrieveSolutions(string solutionName)
        {
            QueryExpression query = new QueryExpression
            {
                EntityName = "solution",
                ColumnSet = new ColumnSet(true),
                Criteria =
                {
                    Conditions =
                    {
                        new ConditionExpression("uniquename", ConditionOperator.Equal, solutionName)
                    },
                }
            };

            RetrieveMultipleRequest request = new RetrieveMultipleRequest();
            request.Query = query;

            try
            {
                RetrieveMultipleResponse response = (RetrieveMultipleResponse)service.Execute(request);
                EntityCollection results = response.EntityCollection;
                return results;
            }
            catch (FaultException<OrganizationServiceFault> ex)
            {
                List<KeyValuePair<string, string>> parameters = new List<KeyValuePair<string, string>>()
                {
                    new KeyValuePair<string, string>("uniquename", solutionName)
                };

                CRMCommandException exc = new CRMCommandException();
                ex.Data.Add("CommandName", "Execute");
                ex.Data.Add("MethodName", "RetrieveSolutions");
                ex.Data.Add("Parameters", parameters);

                throw new Exception(ex.Message, exc);
            }
        }

        public bool ImportSolution(string fileName)
        {
            byte[] fileBytes = File.ReadAllBytes(fileName);

            ImportSolutionRequest request = new ImportSolutionRequest()
            {
                CustomizationFile = fileBytes
            };

            try
            {
                ImportSolutionResponse response = (ImportSolutionResponse)service.Execute(request);
                return true;
            }
            catch (FaultException<OrganizationServiceFault> ex)
            {
                List<KeyValuePair<string, string>> parameters = new List<KeyValuePair<string, string>>()
                {
                    
                };

                CRMCommandException exc = new CRMCommandException();
                ex.Data.Add("CommandName", "Execute");
                ex.Data.Add("MethodName", "ImportSolution");
                ex.Data.Add("Parameters", parameters);

                throw new Exception(ex.Message, exc);
            }
        }

        public Guid RetrieveRole(string roleName)
        {
            QueryExpression query = new QueryExpression
            {
                EntityName = "role",
                ColumnSet = new ColumnSet("roleid"),
                Criteria = new FilterExpression
                {
                    Conditions =
                    {

                        new ConditionExpression
                        {
                            AttributeName = "name",
                            Operator = ConditionOperator.Equal,
                            Values = {roleName}
                        }
                    }
                }
            };

            Guid roleId = Guid.Empty;
            try
            {
                EntityCollection roles = service.RetrieveMultiple(query);

                if (roles.Entities.Count > 0)
                {
                    Entity exchangeSyncRole = service.RetrieveMultiple(query).Entities[0];
                    roleId = exchangeSyncRole.Id;
                }
            }
            catch (FaultException<OrganizationServiceFault> ex)
            {
                // Logger.AddLog(EventLevel.Error, DateTime.Now, "RetrieveRole", "CRM", "Error Retrieving Exchange Sync Security Role for CRM", ex.Message);
                List<KeyValuePair<string, string>> parameters = new List<KeyValuePair<string, string>>()
                {
                    new KeyValuePair<string, string>("name", roleName)
                };

                CRMCommandException exc = new CRMCommandException();
                ex.Data.Add("CommandName", "RetrieveMultiple");
                ex.Data.Add("MethodName", "RetrieveRole");
                ex.Data.Add("Parameters", parameters);

                throw new Exception(ex.Message, exc);
            }
            return roleId;
        }

        public void AssociateUserToRole(Guid userId, Guid roleId)
        {
            if ((roleId != Guid.Empty) && (userId != Guid.Empty))
            {
                try
                {
                    service.Associate("systemuser", userId, new Relationship("systemuserroles_association"), new EntityReferenceCollection() { new EntityReference("role", roleId) });
                }
                catch (FaultException<OrganizationServiceFault> ex)
                {
                    List<KeyValuePair<string, string>> parameters = new List<KeyValuePair<string, string>>()
                    {
                    
                    };

                    CRMCommandException exc = new CRMCommandException();
                    ex.Data.Add("CommandName", "Associate");
                    ex.Data.Add("MethodName", "AssociateUserToRole");
                    ex.Data.Add("Parameters", parameters);

                    throw new Exception(ex.Message, exc);
                    // Logger.AddLog(EventLevel.Error, DateTime.Now, "AssociateUserToRole", "CRM", "Error Associated User to Exchange Sync Security Role", ex.Message);
                }

            }
        }


        #endregion

        #region Application Settings Functions

        public Guid CreateApplicationSetting(int categoryCode, string key, string value, string description = "")
        {
            Entity appSetting = new Entity("xrm_appsettings");
            appSetting.Attributes["xrm_key"] = key;
            appSetting.Attributes["xrm_value"] = value;
            appSetting.Attributes["xrm_category"] = new OptionSetValue(categoryCode);

            if (!string.IsNullOrEmpty(description))
                appSetting.Attributes["xrm_description"] = description;

            try
            {
                Guid settingId = service.Create(appSetting);
                return settingId;
            }
            catch (FaultException<OrganizationServiceFault> ex)
            {
                // Logger.AddLog(EventLevel.Error, DateTime.Now, "CreateApplicationSetting", "CRM", String.Format("Error Calling UpdateApplicationSetting for key {0} ", key), ex.Message);
                List<KeyValuePair<string, string>> parameters = new List<KeyValuePair<string, string>>()
                {
                    new KeyValuePair<string, string>("xrm_key", key),
                    new KeyValuePair<string, string>("xrm_value", value),
                    new KeyValuePair<string, string>("xrm_category", categoryCode.ToString()),
                    new KeyValuePair<string, string>("xrm_description", description)
                };

                CRMCommandException exc = new CRMCommandException();
                ex.Data.Add("CommandName", "Create");
                ex.Data.Add("MethodName", "CreateApplicationSetting");
                ex.Data.Add("Parameters", parameters);

                throw new Exception(ex.Message, exc);
            }
        }

        public Guid CreateSyncList(string profileName, string entityName, string entityDisplayName, Guid entityId, string listName, string exchangeGroupName, int groupType = 1, int objectType = 2)
        {
            Entity syncList = new Entity("xrm_synclist");
            syncList.Attributes["xrm_profilename"] = profileName;
            syncList.Attributes["xrm_entityname"] = entityName;
            syncList.Attributes["xrm_entitydisplayname"] = entityDisplayName;
            syncList.Attributes["xrm_entityid"] = entityId.ToString();
            syncList.Attributes["xrm_listname"] = listName;
            syncList.Attributes["xrm_exchangegroupname"] = exchangeGroupName;
            syncList.Attributes["xrm_exchangegrouptypecode"] = new OptionSetValue(groupType);
            syncList.Attributes["xrm_objecttypecode"] = new OptionSetValue(objectType);
            syncList.Attributes["xrm_liststatus"] = true;
            syncList.Attributes["xrm_autosyncstatus"] = true;
            try
            {
                Guid settingId = service.Create(syncList);
                return settingId;
            }
            catch (FaultException<OrganizationServiceFault> ex)
            {
                // Logger.AddLog(EventLevel.Error, DateTime.Now, "CreateApplicationSetting", "CRM", String.Format("Error Calling UpdateApplicationSetting for key {0} ", key), ex.Message);
                List<KeyValuePair<string, string>> parameters = new List<KeyValuePair<string, string>>()
                {
                    new KeyValuePair<string, string>("xrm_profilename", profileName),
                    new KeyValuePair<string, string>("xrm_entityname", entityName),
                    new KeyValuePair<string, string>("xrm_entitydisplayname", entityDisplayName),
                    new KeyValuePair<string, string>("xrm_entityid", entityId.ToString()),
                    new KeyValuePair<string, string>("xrm_listname", listName),
                    new KeyValuePair<string, string>("xrm_exchangegroupname", exchangeGroupName),
                    new KeyValuePair<string, string>("xrm_liststatus", true.ToString()),
                    new KeyValuePair<string, string>("xrm_autosyncstatus", true.ToString())
                };

                CRMCommandException exc = new CRMCommandException();
                ex.Data.Add("CommandName", "Create");
                ex.Data.Add("MethodName", "CreateSyncList");
                ex.Data.Add("Parameters", parameters);

                throw new Exception(ex.Message, exc);
            }

        }

        public void UpdateSyncList(Guid syncListId, string exchangeGroupName, int groupType = 1)
        {
            Entity syncList = new Entity("xrm_synclist");
            syncList.Id = syncListId;
            syncList.Attributes["xrm_exchangegroupname"] = exchangeGroupName;
            syncList.Attributes["xrm_exchangegrouptypecode"] = new OptionSetValue(groupType);


            try
            {
                service.Update(syncList);
            }
            catch (FaultException<OrganizationServiceFault> ex)
            {
                List<KeyValuePair<string, string>> parameters = new List<KeyValuePair<string, string>>()
                {
                    new KeyValuePair<string, string>("xrm_exchangegroupname", exchangeGroupName)
                };

                CRMCommandException exc = new CRMCommandException();
                ex.Data.Add("CommandName", "Update");
                ex.Data.Add("MethodName", "UpdateApplicationSetting");
                ex.Data.Add("Parameters", parameters);

                throw new Exception(ex.Message, exc);
                // Logger.AddLog(EventLevel.Error, DateTime.Now, "UpdateApplicationSetting", "CRM", String.Format("Error Calling UpdateApplicationSetting for id {0} ", keyId.ToString()), ex.Message);
            }
        }

        public void UpdateApplicationSetting(Guid keyId, string value)
        {
            Entity appSetting = new Entity("xrm_appsettings");
            appSetting.Id = keyId;
            appSetting.Attributes["xrm_value"] = value;

            try
            {
                service.Update(appSetting);
            }
            catch (FaultException<OrganizationServiceFault> ex)
            {
                List<KeyValuePair<string, string>> parameters = new List<KeyValuePair<string, string>>()
                {
                    new KeyValuePair<string, string>("xrm_value", value)
                };

                CRMCommandException exc = new CRMCommandException();
                ex.Data.Add("CommandName", "Update");
                ex.Data.Add("MethodName", "UpdateApplicationSetting");
                ex.Data.Add("Parameters", parameters);

                throw new Exception(ex.Message, exc);
                // Logger.AddLog(EventLevel.Error, DateTime.Now, "UpdateApplicationSetting", "CRM", String.Format("Error Calling UpdateApplicationSetting for id {0} ", keyId.ToString()), ex.Message);
            }
        }

        public EntityCollection RetrieveApplicationSettings(string prefix)
        {
            QueryExpression query = new QueryExpression
            {
                ColumnSet = new ColumnSet(new string[] { "xrm_key", "xrm_value", "xrm_securevalue" }),
                EntityName = "xrm_appsettings",
                Criteria =
                {
                    Conditions =
                    {
                        new ConditionExpression("xrm_key", ConditionOperator.BeginsWith, prefix)
                    },
                }
            };

            RetrieveMultipleRequest request = new RetrieveMultipleRequest();
            request.Query = query;

            try
            {

                RetrieveMultipleResponse response = (RetrieveMultipleResponse)service.Execute(request);
                EntityCollection results = response.EntityCollection;
                return results;
            }
            catch (FaultException<OrganizationServiceFault> ex)
            {
                List<KeyValuePair<string, string>> parameters = new List<KeyValuePair<string, string>>()
                {
                    new KeyValuePair<string, string>("xrm_key", prefix)
                };

                CRMCommandException exc = new CRMCommandException();
                ex.Data.Add("CommandName", "RetrieveMultipleRequest");
                ex.Data.Add("MethodName", "RetrieveApplicationSettings");
                ex.Data.Add("Parameters", parameters);

                throw new Exception(ex.Message, exc);
            }
        }

        public EntityCollection RetrieveApplicationSettings(int categoryCode)
        {
            QueryExpression query = new QueryExpression
            {
                ColumnSet = new ColumnSet(new string[] { "xrm_key", "xrm_value", "xrm_securevalue" }),
                EntityName = "xrm_appsettings",
                Criteria =
                {
                    Conditions =
                    {
                        new ConditionExpression("xrm_category", ConditionOperator.Equal, categoryCode)
                    },
                }
            };

            RetrieveMultipleRequest request = new RetrieveMultipleRequest();
            request.Query = query;

            try
            {
                
                RetrieveMultipleResponse response = (RetrieveMultipleResponse)service.Execute(request);
                EntityCollection results = response.EntityCollection;
                return results;
            }
            catch (FaultException<OrganizationServiceFault> ex)
            {
                List<KeyValuePair<string, string>> parameters = new List<KeyValuePair<string, string>>()
                {
                    new KeyValuePair<string, string>("xrm_category", categoryCode.ToString())
                };

                CRMCommandException exc = new CRMCommandException();
                ex.Data.Add("CommandName", "RetrieveMultipleRequest");
                ex.Data.Add("MethodName", "RetrieveApplicationSettings");
                ex.Data.Add("Parameters", parameters);

                throw new Exception(ex.Message, exc);
            }
        }

        public Entity RetrieveApplicationSetting(Guid appSettingId)
        {
            ColumnSet columns = new ColumnSet(new string[] { "xrm_key", "xrm_value", "xrm_securevalue" });
            Entity appSetting = service.Retrieve("xrm_appsettings", appSettingId, columns);
            return appSetting;
        }

        public Entity RetrieveApplicationSetting(string key)
        {
            QueryExpression query = new QueryExpression
            {
                ColumnSet = new ColumnSet(new string[] { "xrm_key", "xrm_value", "xrm_securevalue" }),
                EntityName = "xrm_appsettings",
                Criteria =
                {
                    Conditions =
                    {
                        new ConditionExpression("xrm_key", ConditionOperator.Equal, key)
                    }
                }
            };

            try
            {
                EntityCollection results = service.RetrieveMultiple(query);
                if (results.Entities.Count > 0)
                {
                    return results.Entities[0];
                }
                else
                    return null;
            }
            catch (FaultException<OrganizationServiceFault> ex)
            {
                List<KeyValuePair<string, string>> parameters = new List<KeyValuePair<string, string>>()
                {
                    new KeyValuePair<string, string>("xrm_key", key)
                };

                CRMCommandException exc = new CRMCommandException();
                ex.Data.Add("CommandName", "RetrieveMultipleRequest");
                ex.Data.Add("MethodName", "RetrieveApplicationSetting");
                ex.Data.Add("Parameters", parameters);

                throw new Exception(ex.Message, exc);
            }

        }

        #endregion

        #region Field Mapping Functions

        public Guid CreateFieldMapping(string crmFieldName, string exchangeFieldName, int fieldType, int dependencyType)
        {
            Entity fieldMapping = new Entity("xrm_fieldmapping");
            fieldMapping.Attributes["xrm_fieldname"] = crmFieldName;
            fieldMapping.Attributes["xrm_exchangefieldname"] = exchangeFieldName;
            fieldMapping.Attributes["xrm_type"] = new OptionSetValue(fieldType);
            fieldMapping.Attributes["xrm_dependency"] = new OptionSetValue(dependencyType);

            try
            {
                Guid settingId = service.Create(fieldMapping);
                return settingId;
            }
            catch (FaultException<OrganizationServiceFault> ex)
            {
                // Logger.AddLog(EventLevel.Error, DateTime.Now, "CreateApplicationSetting", "CRM", String.Format("Error Calling UpdateApplicationSetting for key {0} ", key), ex.Message);
                List<KeyValuePair<string, string>> parameters = new List<KeyValuePair<string, string>>()
                {
                    new KeyValuePair<string, string>("xrm_fieldname", crmFieldName),
                    new KeyValuePair<string, string>("xrm_exchangefieldname", exchangeFieldName),
                    new KeyValuePair<string, string>("xrm_type", fieldType.ToString()),
                    new KeyValuePair<string, string>("xrm_dependency", dependencyType.ToString())
                };

                CRMCommandException exc = new CRMCommandException();
                ex.Data.Add("CommandName", "Create");
                ex.Data.Add("MethodName", "CreateFieldMapping");
                ex.Data.Add("Parameters", parameters);

                throw new Exception(ex.Message, exc);
            }

        }

        public EntityCollection RetrieveFieldMap()
        {
            QueryExpression query = new QueryExpression
            {
                ColumnSet = new ColumnSet(new string[] { "xrm_fieldname", "xrm_type", "xrm_exchangefieldname" }),
                EntityName = "xrm_fieldmapping",
            };

            RetrieveMultipleRequest request = new RetrieveMultipleRequest();
            request.Query = query;

            try
            {

                RetrieveMultipleResponse response = (RetrieveMultipleResponse)service.Execute(request);
                EntityCollection results = response.EntityCollection;
                return results;
            }
            catch (FaultException<OrganizationServiceFault> ex)
            {
                List<KeyValuePair<string, string>> parameters = new List<KeyValuePair<string, string>>()
                {
                };

                CRMCommandException exc = new CRMCommandException();
                ex.Data.Add("CommandName", "RetrieveMultipleRequest");
                ex.Data.Add("MethodName", "RetrieveFieldMap");
                ex.Data.Add("Parameters", parameters);

                throw new Exception(ex.Message, exc);
            }
        }

        public Entity RetrieveFieldMap(string crmFieldName)
        {
            QueryExpression query = new QueryExpression
            {
                ColumnSet = new ColumnSet(new string[] { "xrm_fieldmappingid", "xrm_fieldname" }),
                EntityName = "xrm_fieldmapping",
                Criteria =
                {
                    Conditions =
                    {
                        new ConditionExpression("xrm_fieldname", ConditionOperator.Equal, crmFieldName)
                    },
                }
            };


            try
            {
                EntityCollection results = service.RetrieveMultiple(query);
                if (results.Entities.Count > 0)
                {
                    return results.Entities[0];
                }
                else
                    return null;
            }
            catch (FaultException<OrganizationServiceFault> ex)
            {
                List<KeyValuePair<string, string>> parameters = new List<KeyValuePair<string, string>>()
                {
                    new KeyValuePair<string, string>("xrm_fieldname", crmFieldName)
                };

                CRMCommandException exc = new CRMCommandException();
                ex.Data.Add("CommandName", "RetrieveMultipleRequest");
                ex.Data.Add("MethodName", "RetrieveFieldMap");
                ex.Data.Add("Parameters", parameters);

                throw new Exception(ex.Message, exc);
            }

        }


        #endregion

        #region User Job Functions

        public Guid CreateUserJob(Guid jobRuleId, int jobType, string jobName, DateTime actualStart, DateTime actualEnd, string logfilename, int errorCount)
        {
            Guid userjobid = Guid.Empty;
            Entity userjob = new Entity("xrm_userjob");
            if (jobRuleId != Guid.Empty)
                userjob.Attributes["xrm_userjobruleid"] = new EntityReference("xrm_userjobrule", jobRuleId);

            userjob.Attributes["xrm_jobtype"] = new OptionSetValue(jobType);
            userjob.Attributes["xrm_name"] = jobName;
            userjob.Attributes["xrm_actualstart"] = actualStart;
            userjob.Attributes["xrm_actualend"] = actualEnd;
            userjob.Attributes["xrm_errorcount"] = errorCount;
            userjob.Attributes["xrm_logfilename"] = logfilename;
            bool success = true;
            if (errorCount > 0)
                success = false;

            userjob.Attributes["xrm_completionstatus"] = success;
            try
            {
                userjobid = service.Create(userjob);
            }
            catch (FaultException<OrganizationServiceFault> ex)
            {
                List<KeyValuePair<string, string>> parameters = new List<KeyValuePair<string, string>>()
                {
                    new KeyValuePair<string, string>("xrm_userjobruleid", jobRuleId.ToString()),
                    new KeyValuePair<string, string>("xrm_jobtype", jobType.ToString()),
                    new KeyValuePair<string, string>("xrm_name", jobName),
                    new KeyValuePair<string, string>("xrm_actualstart", actualStart.ToString()),
                    new KeyValuePair<string, string>("xrm_actualend", actualEnd.ToString()),
                    new KeyValuePair<string, string>("xrm_errorcount", errorCount.ToString()),
                    new KeyValuePair<string, string>("xrm_logfilename", logfilename),
                    new KeyValuePair<string, string>("xrm_completionstatus", success.ToString())
                };

                CRMCommandException exc = new CRMCommandException();
                ex.Data.Add("CommandName", "Create");
                ex.Data.Add("MethodName", "CreateUserJob");
                ex.Data.Add("Parameters", parameters);

                throw new Exception(ex.Message, exc);
            }
            return userjobid;
        }

        public List<KeyValuePair<string, string>> RetrieveAllEntities()
        {
            List<KeyValuePair<string, string>> entities = new List<KeyValuePair<string,string>>();
 
            RetrieveAllEntitiesRequest request = new RetrieveAllEntitiesRequest()
            {
                EntityFilters = Microsoft.Xrm.Sdk.Metadata.EntityFilters.Entity,
                RetrieveAsIfPublished = true,
            };

            try
            {
                RetrieveAllEntitiesResponse response = (RetrieveAllEntitiesResponse)service.Execute(request);
                if (response.EntityMetadata.Length > 0)
                {
                    foreach (EntityMetadata entity in response.EntityMetadata)
                    {
                        string entityName = entity.LogicalName;
                        string displayName = entity.DisplayName.UserLocalizedLabel.Label;
                        string primaryAttribute = entity.PrimaryNameAttribute;
                        entities.Add(new KeyValuePair<string,string>(entityName, displayName));
                    }
                }
            }
            catch (FaultException<OrganizationServiceFault> ex)
            {
                List<KeyValuePair<string, string>> parameters = new List<KeyValuePair<string, string>>()
                {
                };

                CRMCommandException exc = new CRMCommandException();
                ex.Data.Add("CommandName", "Execute");
                ex.Data.Add("MethodName", "RetrieveAllEntities");
                ex.Data.Add("Parameters", parameters);

                throw new Exception(ex.Message, exc);
            }
            return entities;
            
        }

        #endregion

        #region Deprecated User Job Functions

        public EntityCollection RetrieveDailyPendingJobs(int timeofday, PeriodOfDay periodofday, DayOfWeek weekday, string profileName = "")
        {
            EntityCollection entities = new EntityCollection();

            // Retrieve Daily
            string[] columnsArray = new string[] { "xrm_userjobruleid", "xrm_name", "xrm_jobtype" };
            ColumnSet columns = new ColumnSet(columnsArray);
            QueryExpression query = new QueryExpression
            {
                ColumnSet = columns,
                EntityName = "xrm_userjobrule",
                Orders =
                {
                    new OrderExpression("xrm_jobtype", OrderType.Ascending)
                }
            };

            // Sunday - Only Every Day
            // Monday - Every day, Every weekday
            // Tuesday - Friday - everyday, everyweekday, every post weekday
            // Saturday - Every day, every post weekday

            FilterExpression filter1 = new FilterExpression(LogicalOperator.Or);
            switch (weekday)
            {
                case DayOfWeek.Sunday:
                    filter1.Conditions.Add(new ConditionExpression("xrm_dailyfrequencytype", ConditionOperator.Equal, 7));
                    break;
                case DayOfWeek.Monday:
                    filter1.Conditions.Add(new ConditionExpression("xrm_dailyfrequencytype", ConditionOperator.Equal, 7));
                    filter1.Conditions.Add(new ConditionExpression("xrm_dailyfrequencytype", ConditionOperator.Equal, 5));
                    break;
                case DayOfWeek.Saturday:
                    filter1.Conditions.Add(new ConditionExpression("xrm_dailyfrequencytype", ConditionOperator.Equal, 7));
                    filter1.Conditions.Add(new ConditionExpression("xrm_dailyfrequencytype", ConditionOperator.Equal, 6));
                    break;
                default:
                    filter1.Conditions.Add(new ConditionExpression("xrm_dailyfrequencytype", ConditionOperator.Equal, 7));
                    filter1.Conditions.Add(new ConditionExpression("xrm_dailyfrequencytype", ConditionOperator.Equal, 5));
                    filter1.Conditions.Add(new ConditionExpression("xrm_dailyfrequencytype", ConditionOperator.Equal, 6));
                    break;
            }

            FilterExpression filter2 = new FilterExpression(LogicalOperator.And);
            filter2.Conditions.Add(new ConditionExpression("xrm_timeofday", ConditionOperator.Equal, timeofday));
            filter2.Conditions.Add(new ConditionExpression("xrm_periodofday", ConditionOperator.Equal, (int)periodofday));

            FilterExpression filter3 = new FilterExpression(LogicalOperator.And);
            filter3.Conditions.Add(new ConditionExpression("statecode", ConditionOperator.Equal, 0));

            FilterExpression filter = new FilterExpression(LogicalOperator.And);
            filter.AddFilter(filter1); filter.AddFilter(filter2); filter.AddFilter(filter3);

            if (!string.IsNullOrEmpty(profileName))
            {
                FilterExpression filter4 = new FilterExpression(LogicalOperator.And);
                filter4.Conditions.Add(new ConditionExpression("xrm_profilename", ConditionOperator.Equal, profileName));
                filter.AddFilter(filter4);
            }

            query.Criteria = filter;

            try
            {
                entities = service.RetrieveMultiple(query);
            }
            catch (FaultException<OrganizationServiceFault> ex)
            {
                List<KeyValuePair<string, string>> parameters = new List<KeyValuePair<string, string>>()
                {
                    new KeyValuePair<string, string>("xrm_dailyfrequencytype", weekday.ToPlainString())
                };

                CRMCommandException exc = new CRMCommandException();
                ex.Data.Add("CommandName", "RetrieveMultiple");
                ex.Data.Add("MethodName", "RetrieveDailyPendingJobs");
                ex.Data.Add("Parameters", parameters);

                throw new Exception(ex.Message, exc);
            }
            return entities;
        }

        public EntityCollection RetrieveWeeklyPendingJobs(int timeofday, PeriodOfDay periodofday, DayOfWeek weekday, string profileName = "")
        {
            EntityCollection entities = new EntityCollection();

            // Retrieve Daily
            string[] columnsArray = new string[] { "xrm_userjobruleid", "xrm_name", "xrm_jobtype" };
            ColumnSet columns = new ColumnSet(columnsArray);
            QueryExpression query = new QueryExpression
            {
                ColumnSet = columns,
                EntityName = "xrm_userjobrule",
                Orders =
                {
                    new OrderExpression("xrm_jobtype", OrderType.Ascending)
                }
            };

            int dayofweek = (int)weekday + 1;
            FilterExpression filter1 = new FilterExpression(LogicalOperator.And);
            filter1.Conditions.Add(new ConditionExpression("xrm_dayofweek", ConditionOperator.Equal, dayofweek));
            filter1.Conditions.Add(new ConditionExpression("xrm_timeofday", ConditionOperator.Equal, timeofday));
            filter1.Conditions.Add(new ConditionExpression("xrm_periodofday", ConditionOperator.Equal, (int)periodofday));

            FilterExpression filter2 = new FilterExpression(LogicalOperator.And);
            filter2.Conditions.Add(new ConditionExpression("statecode", ConditionOperator.Equal, 0));

            FilterExpression filter = new FilterExpression(LogicalOperator.And);
            filter.AddFilter(filter1); filter.AddFilter(filter2);

            if (!string.IsNullOrEmpty(profileName))
            {
                FilterExpression filter3 = new FilterExpression(LogicalOperator.And);
                filter3.Conditions.Add(new ConditionExpression("xrm_profilename", ConditionOperator.Equal, profileName));
                filter.AddFilter(filter3);
            }

            query.Criteria = filter;


            try
            {
                entities = service.RetrieveMultiple(query);
            }
            catch (FaultException<OrganizationServiceFault> ex)
            {
                List<KeyValuePair<string, string>> parameters = new List<KeyValuePair<string, string>>()
                {
                    new KeyValuePair<string, string>("xrm_dayofweek", dayofweek.ToString()),
                    new KeyValuePair<string, string>("xrm_timeofday", timeofday.ToString()),
                    new KeyValuePair<string, string>("xrm_periodofday", periodofday.ToPlainString()),
                    new KeyValuePair<string, string>("statecode", "0")
                };

                CRMCommandException exc = new CRMCommandException();
                ex.Data.Add("CommandName", "RetrieveMultiple");
                ex.Data.Add("MethodName", "RetrieveWeeklyPendingJobs");
                ex.Data.Add("Parameters", parameters);

                throw new Exception(ex.Message, exc);
            }
            return entities;
        }

        public EntityCollection RetrieveMonthlyPendingJobs(int timeofday, PeriodOfDay periodofday, int dayofmonth, bool islastday, string profileName = "")
        {
            EntityCollection entities = new EntityCollection();

            // Retrieve Daily
            string[] columnsArray = new string[] { "xrm_userjobruleid", "xrm_name", "xrm_jobtype" };
            ColumnSet columns = new ColumnSet(columnsArray);
            QueryExpression query = new QueryExpression
            {
                ColumnSet = columns,
                EntityName = "xrm_userjobrule",
                Orders =
                {
                    new OrderExpression("xrm_jobtype", OrderType.Ascending)
                }
            };

            FilterExpression filter1 = new FilterExpression(LogicalOperator.And);
            if (islastday)
                filter1.Conditions.Add(new ConditionExpression("xrm_dayofmonth", ConditionOperator.GreaterEqual, dayofmonth));
            else
                filter1.Conditions.Add(new ConditionExpression("xrm_dayofmonth", ConditionOperator.Equal, dayofmonth));
            filter1.Conditions.Add(new ConditionExpression("xrm_timeofday", ConditionOperator.Equal, timeofday));
            filter1.Conditions.Add(new ConditionExpression("xrm_periodofday", ConditionOperator.Equal, (int)periodofday));

            FilterExpression filter2 = new FilterExpression(LogicalOperator.And);
            filter2.Conditions.Add(new ConditionExpression("statecode", ConditionOperator.Equal, 0));

            FilterExpression filter = new FilterExpression(LogicalOperator.And);
            filter.AddFilter(filter1); filter.AddFilter(filter2);

            if (!string.IsNullOrEmpty(profileName))
            {
                FilterExpression filter3 = new FilterExpression(LogicalOperator.And);
                filter3.Conditions.Add(new ConditionExpression("xrm_profilename", ConditionOperator.Equal, profileName));
                filter.AddFilter(filter3);
            }

            query.Criteria = filter;

            try
            {
                entities = service.RetrieveMultiple(query);
            }
            catch (FaultException<OrganizationServiceFault> ex)
            {
                List<KeyValuePair<string, string>> parameters = new List<KeyValuePair<string, string>>()
                {
                    new KeyValuePair<string, string>("xrm_dayofmonth", dayofmonth.ToString()),
                    new KeyValuePair<string, string>("xrm_timeofday", timeofday.ToString()),
                    new KeyValuePair<string, string>("xrm_periodofday", periodofday.ToPlainString()),
                    new KeyValuePair<string, string>("statecode", "0")
                };

                CRMCommandException exc = new CRMCommandException();
                ex.Data.Add("CommandName", "RetrieveMultiple");
                ex.Data.Add("MethodName", "RetrieveMonthlyPendingJobs");
                ex.Data.Add("Parameters", parameters);

                throw new Exception(ex.Message, exc);
            }
            return entities;

        }


        #endregion

    }


    public class CRMEntity
    {
        public string EntityName { get; set; }
        public string DisplayName { get; set; }
        public string PrimaryAttributeName { get; set; }

        public CRMEntity() { }

        public void AddEntity(string entityName, string displayName, string primaryAttributeName)
        {
            this.EntityName = entityName;
            this.DisplayName = displayName;
            this.PrimaryAttributeName = primaryAttributeName;
        }
    }

}
