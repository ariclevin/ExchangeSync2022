using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using Microsoft.Xrm.Sdk;

namespace ExchangeSync
{
    public partial class Core
    {
        public EntityCollection RetrieveContacts(string entityName, string listName, Guid listId)
        {
            EntityCollection contacts = new EntityCollection();
            if (entityName == "list")
            {
                bool listType = false;
                string query = string.Empty;
                try
                {
                    Entity list = crm.RetrieveList(listId);
                    listType = (bool)list.Attributes["type"];
                    query = list.Contains("query") ? list.Attributes["query"].ToString() : string.Empty;
                }
                catch (System.Exception ex)
                {
                    string message = ex.Message;
                    string command = "", function = "", parameterString = "";
                    Helper.LogException(ex, ref command, ref function, ref parameterString);
                    Trace.AddLog(EventLevel.Error, DateTime.Now, "CRM", function, command, message, string.Empty, parameterString);
                }

                if (listType == false)
                {
                    // Processing Static List
                    Trace.AddLog(EventLevel.Information, DateTime.Now, "RetrieveContacts", "CRM", "Processing Static List: " + listName, string.Empty);
                    try
                    {
                        contacts = crm.RetrieveContactListMembers(listId);
                        if (contacts.Entities.Count > 0)
                            Trace.AddLog(EventLevel.Information, DateTime.Now, "RetrieveContacts", "CRM", String.Format("{0} Contacts retrieved from Static List {1} ", contacts.Entities.Count, listName), string.Empty);
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
                else
                {
                    // Processing Dynamic List
                    Trace.AddLog(EventLevel.Information, DateTime.Now, "RetrieveContacts", "CRM", "Processing Dynamic List: " + listName, string.Empty);
                    try
                    {
                        contacts = crm.RetrieveListMembers(query, true);
                        if (contacts.Entities.Count > 0)
                            Trace.AddLog(EventLevel.Information, DateTime.Now, "RetrieveContacts", "CRM", String.Format("{0} Contacts retrieved from Dynamic List {1} ", contacts.Entities.Count, listName), string.Empty);

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
            }
            else
            {
                Trace.AddLog(EventLevel.Information, DateTime.Now, "RetrieveContacts", "CRM", "Processing Custom Entity: " + listName, string.Empty);
                try
                {
                    contacts = RetrieveCustomEntityInfo(entityName, listId);
                    if (contacts.Entities.Count > 0)
                        Trace.AddLog(EventLevel.Information, DateTime.Now, "RetrieveContacts", "CRM", String.Format("{0} Contacts retrieved from Custom Entity {1} ", contacts.Entities.Count, entityName), string.Empty);
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

            return contacts;
        }

        public EntityCollection RetrieveLeads(string entityName, string listName, Guid listId)
        {
            EntityCollection leads = new EntityCollection();
            if (entityName == "list")
            {
                bool listType = false;
                string query = string.Empty;
                try
                {
                    Entity list = crm.RetrieveList(listId);
                    listType = (bool)list.Attributes["type"];
                    query = list.Contains("query") ? list.Attributes["query"].ToString() : string.Empty;
                }
                catch (System.Exception ex)
                {
                    string message = ex.Message;
                    string command = "", function = "", parameterString = "";
                    Helper.LogException(ex, ref command, ref function, ref parameterString);
                    Trace.AddLog(EventLevel.Error, DateTime.Now, "CRM", function, command, message, string.Empty, parameterString);
                }

                if (listType == false)
                {
                    // Processing Static List
                    Trace.AddLog(EventLevel.Information, DateTime.Now, "RetrieveLeads", "CRM", "Processing Static List: " + listName, string.Empty);
                    try
                    {
                        leads = crm.RetrieveLeadListMembers(listId);
                        if (leads.Entities.Count > 0)
                            Trace.AddLog(EventLevel.Information, DateTime.Now, "RetrieveLeads", "CRM", String.Format("{0} Leads retrieved from Static List {1} ", leads.Entities.Count, listName), string.Empty);
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
                else
                {
                    // Processing Dynamic List
                    Trace.AddLog(EventLevel.Information, DateTime.Now, "RetrieveLeads", "CRM", "Processing Dynamic List: " + listName, string.Empty);
                    try
                    {
                        leads = crm.RetrieveListMembers(query, true);
                        if (leads.Entities.Count > 0)
                            Trace.AddLog(EventLevel.Information, DateTime.Now, "RetrieveLeads", "CRM", String.Format("{0} Leads retrieved from Dynamic List {1} ", leads.Entities.Count, listName), string.Empty);

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
            }
            else
            {
                Trace.AddLog(EventLevel.Information, DateTime.Now, "RetrieveLeads", "CRM", "Processing Custom Entity: " + listName, string.Empty);
                try
                {
                    leads = RetrieveCustomEntityInfo(entityName, listId);
                    if (leads.Entities.Count > 0)
                        Trace.AddLog(EventLevel.Information, DateTime.Now, "RetrieveLeads", "CRM", String.Format("{0} Leads retrieved from Custom Entity {1} ", leads.Entities.Count, entityName), string.Empty);
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

            return leads;
        }


    }
}
