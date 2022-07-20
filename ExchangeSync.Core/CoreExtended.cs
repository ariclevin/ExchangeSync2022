using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

using Microsoft.Xrm.Sdk;

namespace ExchangeSync
{
    public partial class Core
    {
        public void UpdateExchangeContacts(string exchangeListName, string externalDomain, ref string exchangeAlias, ref EntityCollection contacts)
        {
            Trace.AddLog(EventLevel.Verbose, DateTime.Now, "UpdateExchangeContacts", "CRM", String.Format("Map Count = {0}", ProgramSetting.Map.Count.ToString()), "");
            if (ProgramSetting.Map.Count > 0)
            {
                List<string> attributes = new List<string>(); List<string> values = new List<string>();
                foreach (Entity contact in contacts.Entities)
                {
                    // Required Attributes
                    string fullName = contact["fullname"].ToString();
                    string firstName = contact.Contains("firstname") ? contact["firstname"].ToString() : "";
                    string lastName = contact.Contains("lastname") ? contact["lastname"].ToString() : "";
                    string emailAddress = contact.Contains("emailaddress1") ? contact.Attributes["emailaddress1"].ToString() : "";
                    if (string.IsNullOrEmpty(emailAddress))
                        emailAddress = contact.Contains("xrm_preferredemailaddress") ? contact.Attributes["xrm_preferredemailaddress"].ToString() : "";

                    exchangeAlias = contact.Contains(AppSetting.ContactExchangeAliasFieldName.Value) ? contact.Attributes[AppSetting.ContactExchangeAliasFieldName.Value].ToString() : "";
                    int revisionNo = contact.Attributes.ContainsKey(AppSetting.ContactRevisionFieldName.Value) ? Convert.ToInt32(contact.Attributes[AppSetting.ContactRevisionFieldName.Value]) : -1;
                    int statecode = ((OptionSetValue)(contact.Attributes["statecode"])).Value;
                    string custom = "";

                    string customfieldname = AppSetting.ContactAutoNumberFieldName.Value;
                    if (customfieldname.Contains(':'))
                        customfieldname = customfieldname.Substring(0, customfieldname.IndexOf(':'));

                    Trace.AddLog(EventLevel.Verbose, DateTime.Now, "UpdateExchangeContacts", "CRM", String.Format("Custom Field Name = {0}", customfieldname), "");

                    if (!string.IsNullOrEmpty(customfieldname))
                    {
                        // Added 4/24 to Support Custom Entity Values
                        if (contact.Attributes.ContainsKey(customfieldname))
                        {
                            switch (contact.Attributes[customfieldname].GetType().FullName)
                            {
                                case "Microsoft.Xrm.Sdk.EntityReference":
                                    custom = ((EntityReference)(contact.Attributes[customfieldname])).Name;
                                    break;
                                case "System.String":
                                    custom = contact.Attributes[customfieldname].ToString();
                                    break;
                                case "System.Int32":
                                    custom = contact.Attributes[customfieldname].ToString();
                                    break;
                                default:
                                    custom = contact.Attributes[customfieldname].ToString();
                                    break;
                            }
                        }
                        Trace.AddLog(EventLevel.Verbose, DateTime.Now, "UpdateExchangeContacts", "CRM", String.Format("Custom Field Name Value = {0}", custom), "");

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
                        Trace.AddLog(EventLevel.Verbose, DateTime.Now, "UpdateExchangeContacts", "CRM", String.Format("Custom Field Name Value = {0}", custom), "");
                    }

                    if (statecode == 0) // Only for Active Contacts
                    {
                        foreach (FieldMap item in ProgramSetting.Map)
                        {
                            if (contact.Contains(item.CRMFieldName))
                            {
                                string ExchangeFieldName = item.ExchangeFieldName;
                                string ExchangeFieldValue = "";
                                switch (item.CRMFieldType.Key)
                                {
                                    case 1: // Single Line of Text
                                    case 8: // Multiple Lines of Text
                                        switch (ExchangeFieldName)
                                        {
                                            case "DisplayName":
                                                ExchangeFieldValue = ExchangeHelper.GenerateDisplayName(AppSetting.ExchangeMailContactFormat.Value, fullName, firstName, lastName);
                                                break;
                                            case "FirstName":
                                                ExchangeFieldValue = ExchangeHelper.TrimLength(firstName, 64);
                                                break;
                                            case "LastName":
                                                ExchangeFieldValue = ExchangeHelper.TrimLength(lastName, 64);
                                                break;
                                            default:
                                                ExchangeFieldValue = contact.Attributes[item.CRMFieldName].ToString();
                                                break;
                                        }
                                        attributes.Add(ExchangeFieldName); values.Add(ExchangeFieldValue);
                                        break;
                                    case 2: // Option Set
                                        ExchangeFieldValue = contact.FormattedValues[item.CRMFieldName].ToString();
                                        attributes.Add(ExchangeFieldName); values.Add(ExchangeFieldValue);
                                        break;
                                    case 3: // Two Options
                                        break;
                                    case 4: // Whole Number
                                    case 5: // Floating Point Number
                                    case 6: // Decimal Number
                                    case 7: // Currency
                                        ExchangeFieldValue = contact.Attributes[item.CRMFieldName].ToString();
                                        attributes.Add(ExchangeFieldName); values.Add(ExchangeFieldValue);
                                        break;
                                    case 9: // Date and Time
                                        ExchangeFieldValue = contact.Attributes[item.CRMFieldName].ToString();
                                        attributes.Add(ExchangeFieldName); values.Add(ExchangeFieldValue);
                                        break;
                                    case 10: // Lookup
                                        ExchangeFieldValue = ((EntityReference)(contact.Attributes[item.CRMFieldName])).Name;
                                        attributes.Add(ExchangeFieldName); values.Add(ExchangeFieldValue);
                                        break;
                                    default:
                                        ExchangeFieldValue = contact.Attributes[item.CRMFieldName].ToString();
                                        attributes.Add(ExchangeFieldName); values.Add(ExchangeFieldValue);
                                        break;
                                } // end switch
                                Trace.AddLog(EventLevel.Verbose, DateTime.Now, "UpdateExchangeContacts", "CRM", String.Format("CRMFieldName={0}, CRMFieldType={1}, ExchangeFieldName={2}, ExchangeFieldValue={3}", item.CRMFieldName, item.CRMFieldType.Key.ToString(), item.ExchangeFieldName, ExchangeFieldValue), "");

                            } // end if

                        } // end foreach item
                        string[] MailContactAttributes = attributes.ToArray();
                        string[] MailContactValues = values.ToArray();

                        // If this is a Mailbox User call UpdateExchangeUser Method
                        Trace.AddLog(EventLevel.Verbose, DateTime.Now, "UpdateExchangeContacts", "CRM", String.Format("External Domain Name = {0}", externalDomain), "");
                        if (emailAddress.IsEmailInDomain(externalDomain))
                        {
                            UpdateExchangeUser(contact.Id, emailAddress, ref exchangeAlias, revisionNo, fullName, ref MailContactAttributes, ref MailContactValues);
                        }
                        else
                        {

                            // Start Logic for Updating Contact Records Here after fields have been set
                            bool duplicateAlias = false;
                            Trace.AddLog(EventLevel.Verbose, DateTime.Now, "UpdateExchangeContacts", "CRM", String.Format("Exchange Alias = {0}", exchangeAlias), "");
                            if (exchangeAlias.IsValidEmail())
                            {
                                exchangeAlias = ExchangeHelper.GenerateNameFromEmailAddress(exchangeAlias);
                                duplicateAlias = true;
                            }

                            if (!string.IsNullOrEmpty(exchangeAlias))
                            {
                                if (ContactExists(exchangeAlias))
                                {
                                    Trace.AddLog(EventLevel.Information, DateTime.Now, "RunSync", "Exchange", "Calling Update Mail Contact Functions for Alias " + exchangeAlias, string.Empty);
                                    if (duplicateAlias)
                                    {
                                        if (AppSetting.DuplicateSyncAction.Value == "MODIFY")
                                        {
                                            exch.UpdateMailContact(exchangeAlias, exchangeAlias, emailAddress, emailAddress, emailAddress);
                                            UpdateMailContact(exchangeAlias, new string[] { "FirstName", "LastName", "Phone", "Title" }, new string[] { "", "", "", "" });
                                        }
                                    }
                                    else
                                    {
                                        Trace.AddLog(EventLevel.Verbose, DateTime.Now, "UpdateExchangeContacts", "CRM", "Update Mail Contact Functions by Alias", "");
                                        string displayName = ExchangeHelper.GenerateDisplayName(AppSetting.ExchangeMailContactFormat.Value, fullName, firstName, lastName);

                                        UpdateMailContact(exchangeAlias, displayName, emailAddress);
                                        UpdateMailContact(exchangeAlias, MailContactAttributes, MailContactValues);
                                        HideMailContact(exchangeAlias);
                                        System.Threading.Thread.Sleep(100);
                                    }
                                }
                                else if (ContactExists(emailAddress))
                                {
                                    Trace.AddLog(EventLevel.Verbose, DateTime.Now, "UpdateExchangeContacts", "CRM", "CRM Exchange Alias not found, Contact Exists By Email Address?", "");
                                    ExchangeContact current = GetMailContact(emailAddress);
                                    string exchangeAliasFound = current.Alias;

                                    UpdateMailContactAlias(exchangeAliasFound, exchangeAlias); // Updates the Exchange Alias 

                                    Trace.AddLog(EventLevel.Information, DateTime.Now, "RunSync", "Exchange", "Calling Update Mail Contact Functions for Email Address " + emailAddress, string.Empty);

                                    string displayName = ExchangeHelper.GenerateDisplayName(AppSetting.ExchangeMailContactFormat.Value, fullName, firstName, lastName);
                                    UpdateMailContact(exchangeAlias, displayName, emailAddress);

                                    UpdateMailContact(exchangeAlias, MailContactAttributes, MailContactValues);
                                    HideMailContact(exchangeAlias);
                                    System.Threading.Thread.Sleep(100);
                                }
                                else
                                {
                                    if (duplicateAlias)
                                    {
                                        if (AppSetting.DuplicateSyncAction.Value == "MODIFY")
                                        {
                                            try
                                            {
                                                exch.CreateMailContact(emailAddress, AppSetting.ExchangeMailContactOU.Value);
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
                                        }
                                    }
                                    else
                                    {
                                        Trace.AddLog(EventLevel.Verbose, DateTime.Now, "UpdateExchangeContacts", "CRM", "CRM Exchange Alias not found, Create New Mail Contact", "");
                                        Trace.AddLog(EventLevel.Information, DateTime.Now, "RunSync", "Exchange", "Calling Create Mail Contact Functions for " + fullName, string.Empty);
                                        // exch.CreateMailContact(fullName, emailAddress, "", "", AppSetting.ExchangeMailContactOU.Value, customField, out exchangeAlias);
                                        string displayName = ExchangeHelper.GenerateDisplayName(AppSetting.ExchangeMailContactFormat.Value, fullName, firstName, lastName);

                                        try
                                        {
                                            exch.CreateMailContact(displayName, emailAddress, firstName, lastName, AppSetting.ExchangeMailContactOU.Value, custom, out exchangeAlias);
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

                                        UpdateMailContact(exchangeAlias, MailContactAttributes, MailContactValues);
                                        HideMailContact(exchangeAlias);
                                        System.Threading.Thread.Sleep(100);

                                        crm.UpdateContactAlias(contact.Id, exchangeAlias, revisionNo);
                                    }
                                }
                            }
                            else // exchangeAlias IsNullOrEmpty 
                            {
                                Trace.AddLog(EventLevel.Verbose, DateTime.Now, "UpdateExchangeContacts", "CRM", "No Alias in CRM", "");
                                if (!string.IsNullOrEmpty(emailAddress))
                                {
                                    if (ContactExists(emailAddress))
                                    {
                                        Trace.AddLog(EventLevel.Verbose, DateTime.Now, "UpdateExchangeContacts", "CRM", "Contact Exists By Email Address", "");

                                        string exchangeAliasFound = string.Empty;
                                        ExchangeContact current = GetMailContact(emailAddress);
                                        exchangeAliasFound = current.Alias;

                                        string newAlias = ExchangeHelper.FormatAlias(firstName, lastName, custom);
                                        try
                                        {
                                            exch.UpdateMailContactAlias(emailAddress, newAlias);
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

                                        exchangeAlias = newAlias;

                                        Trace.AddLog(EventLevel.Information, DateTime.Now, "RunSync", "Exchange", "Calling Update Mail Contact Functions for " + exchangeAlias, string.Empty);
                                        UpdateMailContact(exchangeAlias, fullName, emailAddress);
                                        UpdateMailContact(exchangeAlias, MailContactAttributes, MailContactValues);
                                        HideMailContact(exchangeAlias);
                                        System.Threading.Thread.Sleep(100);

                                        crm.UpdateContactAlias(contact.Id, exchangeAlias, revisionNo);
                                    }
                                    else
                                    {
                                        Trace.AddLog(EventLevel.Verbose, DateTime.Now, "UpdateExchangeContacts", "CRM", "Contact does not Exist By Email Address", "");
                                        Trace.AddLog(EventLevel.Information, DateTime.Now, "RunSync", "Exchange", "Calling Create Mail Contact Functions for " + fullName, string.Empty);

                                        string displayName = ExchangeHelper.GenerateDisplayName(AppSetting.ExchangeMailContactFormat.Value, fullName, firstName, lastName);
                                        CreateNewContact(emailAddress, displayName, firstName, lastName, custom, ref exchangeAlias);

                                        UpdateMailContact(exchangeAlias, MailContactAttributes, MailContactValues);
                                        HideMailContact(exchangeAlias);
                                        System.Threading.Thread.Sleep(100);

                                        crm.UpdateContactAlias(contact.Id, exchangeAlias, revisionNo);
                                    }
                                }
                            }
                        }

                        if (!string.IsNullOrEmpty(exchangeAlias))
                        {
                            AddDistributionGroupMember(exchangeListName, exchangeAlias);
                            Trace.AddLog(EventLevel.Information, DateTime.Now, "RunSync", "CRM", String.Format("Add Distribution List Member {0} [{1}]", fullName, exchangeAlias), string.Empty);
                        }
                        attributes.Clear(); values.Clear();

                        // TODO: Add Event for increment
                        // tspbStatus.Value++;
                    } // end if statecode == 0
                } // end foreach contact
            }
        }

    }
}
