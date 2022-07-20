using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExchangeSync
{
    // Methods for Connecting to Exchange Server with Error Handling

    public partial class Core
    {
        private bool ContactExists(string identity)
        {
            bool rc = false;
            try
            {
                rc = exch.ContactExists(identity);
            }
            catch (System.Exception ex)
            {
                string message = ex.Message;
                string command = ex.InnerException.Data.Contains("CommandName") ? ex.InnerException.Data["CommandName"].ToString() : string.Empty;
                string function = ex.InnerException.Data.Contains("MethodName") ? ex.InnerException.Data["MethodName"].ToString() : string.Empty;

                List<KeyValuePair<string, string>> parameters = ex.InnerException.Data.Contains("Parameters") ? (List<KeyValuePair<string, string>>)ex.InnerException.Data["Parameters"] : new List<KeyValuePair<string, string>>();
                string parameterString = Helper.KeyValuePairListToString(parameters);
                var eventLevel = EventLevel.Error;
                if (message.Contains("couldn't be found"))
                    eventLevel = EventLevel.Warning;
                Trace.AddLog(eventLevel, DateTime.Now, "Exchange", function, command, message, string.Empty, parameterString);
            }

            return rc;
        }

        private void UpdateMailContact(string identity, string[] exchParams, string[] values)
        {
            bool rc = false;
            try
            {
                rc = exch.UpdateMailContact(identity, exchParams, values);
            }
            catch (System.Exception ex)
            {
                string message = ex.Message;
                string command = ex.InnerException.Data.Contains("CommandName") ? ex.InnerException.Data["CommandName"].ToString() : string.Empty;
                string function = ex.InnerException.Data.Contains("MethodName") ? ex.InnerException.Data["MethodName"].ToString() : string.Empty;

                List<KeyValuePair<string, string>> parameters = ex.InnerException.Data.Contains("Parameters") ? (List<KeyValuePair<string, string>>)ex.InnerException.Data["Parameters"] : new List<KeyValuePair<string, string>>();
                string parameterString = Helper.KeyValuePairListToString(parameters);
                Trace.AddLog(EventLevel.Error, DateTime.Now, "Exchange", function, command, message, string.Empty, parameterString);
            }
        }

        private ExchangeContact GetMailContact(string identity)
        {
            ExchangeContact rc = new ExchangeContact();
            try
            {
                rc = exch.GetMailContact(identity);
            }
            catch (System.Exception ex)
            {
                string message = ex.Message;
                string command = ex.InnerException.Data.Contains("CommandName") ? ex.InnerException.Data["CommandName"].ToString() : string.Empty;
                string function = ex.InnerException.Data.Contains("MethodName") ? ex.InnerException.Data["MethodName"].ToString() : string.Empty;

                List<KeyValuePair<string, string>> parameters = ex.InnerException.Data.Contains("Parameters") ? (List<KeyValuePair<string, string>>)ex.InnerException.Data["Parameters"] : new List<KeyValuePair<string, string>>();
                string parameterString = Helper.KeyValuePairListToString(parameters);
                Trace.AddLog(EventLevel.Error, DateTime.Now, "Exchange", function, command, message, string.Empty, parameterString);

            }
            return rc;

        }

        private bool UpdateMailContact(string identity, string displayName, string emailAddress)
        {
            bool rc = false;
            try
            {
                rc = exch.UpdateMailContact(identity, displayName, emailAddress);
            }
            catch (System.Exception ex)
            {
                string message = ex.Message;
                string command = ex.InnerException.Data.Contains("CommandName") ? ex.InnerException.Data["CommandName"].ToString() : string.Empty;
                string function = ex.InnerException.Data.Contains("MethodName") ? ex.InnerException.Data["MethodName"].ToString() : string.Empty;

                List<KeyValuePair<string, string>> parameters = ex.InnerException.Data.Contains("Parameters") ? (List<KeyValuePair<string, string>>)ex.InnerException.Data["Parameters"] : new List<KeyValuePair<string, string>>();
                string parameterString = Helper.KeyValuePairListToString(parameters);
                Trace.AddLog(EventLevel.Error, DateTime.Now, "Exchange", function, command, message, string.Empty, parameterString);
            }
            return rc;

        }


        private bool UpdateMailContactAlias(string identity, string exchangeAlias)
        {
            bool rc = false;
            try
            {
                rc = exch.UpdateMailContactAlias(identity, exchangeAlias);
            }
            catch (System.Exception ex)
            {
                string message = ex.Message;
                string command = ex.InnerException.Data.Contains("CommandName") ? ex.InnerException.Data["CommandName"].ToString() : string.Empty;
                string function = ex.InnerException.Data.Contains("MethodName") ? ex.InnerException.Data["MethodName"].ToString() : string.Empty;

                List<KeyValuePair<string, string>> parameters = ex.InnerException.Data.Contains("Parameters") ? (List<KeyValuePair<string, string>>)ex.InnerException.Data["Parameters"] : new List<KeyValuePair<string, string>>();
                string parameterString = Helper.KeyValuePairListToString(parameters);
                Trace.AddLog(EventLevel.Error, DateTime.Now, "Exchange", function, command, message, string.Empty, parameterString);
            }
            return rc;

        }

        private bool AddDistributionGroupMember(string exchangeListName, string alias)
        {
            bool rc = false;
            try
            {
                rc = exch.AddDistributionListMember(exchangeListName, alias);
                List<KeyValuePair<string, string>> parameters = new List<KeyValuePair<string, string>>()
                        {
                            new KeyValuePair<string, string>("Identity", exchangeListName),
                            new KeyValuePair<string, string>("Member", alias)
                        };
                string sParams = Helper.KeyValuePairListToString(parameters);
                Trace.AddLog(EventLevel.Information, DateTime.Now, "Exchange", "AddDistributionListMember", "add-distributiongroupmember", "Added Distribution Group Member", string.Empty, sParams);
            }
            catch (System.Exception ex)
            {
                string message = ex.Message;
                string command = ex.InnerException.Data.Contains("CommandName") ? ex.InnerException.Data["CommandName"].ToString() : string.Empty;
                string function = ex.InnerException.Data.Contains("MethodName") ? ex.InnerException.Data["MethodName"].ToString() : string.Empty;

                EventLevel level; // If Already a member of the group display as warning. Added 6/29/2020
                if (message.Contains("is already a member of the group"))
                    level = EventLevel.Warning;
                else
                    level = EventLevel.Error;

                List<KeyValuePair<string, string>> parameters = ex.InnerException.Data.Contains("Parameters") ? (List<KeyValuePair<string, string>>)ex.InnerException.Data["Parameters"] : new List<KeyValuePair<string, string>>();
                string parameterString = Helper.KeyValuePairListToString(parameters);
                Trace.AddLog(level, DateTime.Now, "Exchange", function, command, message, string.Empty, parameterString);
            }

            return rc;

        }

        private void CreateNewContact(string externalEmailAddress, string fullName, string firstName, string lastName, string customField, ref string exchangeAlias, string primarySmtpAddress = "")
        {
            try
            {
                string displayName = ExchangeHelper.GenerateDisplayName(AppSetting.ExchangeMailContactFormat.Value, fullName, firstName, lastName);
                // Add code for inclusion of Primary Smtp Address (Fake Email Address)
                exch.CreateMailContact(displayName, externalEmailAddress, firstName, lastName, AppSetting.ExchangeMailContactOU.Value, customField, out exchangeAlias, primarySmtpAddress);
            }
            catch (System.Exception ex)
            {
                string message = ex.Message;
                string command = ex.InnerException.Data.Contains("CommandName") ? ex.InnerException.Data["CommandName"].ToString() : string.Empty;
                string function = ex.InnerException.Data.Contains("MethodName") ? ex.InnerException.Data["MethodName"].ToString() : string.Empty;

                List<KeyValuePair<string, string>> parameters = ex.InnerException.Data.Contains("Parameters") ? (List<KeyValuePair<string, string>>)ex.InnerException.Data["Parameters"] : new List<KeyValuePair<string, string>>();
                string parameterString = Helper.KeyValuePairListToString(parameters);
                Trace.AddLog(EventLevel.Error, DateTime.Now, "Exchange", function, command, message, string.Empty, parameterString);
            }
        }

        private void HideMailContact(string identity, string value = "")
        {
            if (string.IsNullOrEmpty(value))
                value = AppSetting.HideFromExchangeAddressLists.Value;

            bool? hide = null;
            switch (value)
            {
                case "HIDE":
                    hide = true;
                    //exch.HideMailContact(identity, true);
                    break;
                case "SHOW":
                    hide = false;
                    //exch.HideMailContact(identity, false);
                    break;
                case "IGNORE":
                    break;
                default:
                    break;
            }

            if (hide.HasValue)
            {
                try
                {
                    bool rc = exch.HideMailContact(identity, hide.Value);
                    List<KeyValuePair<string, string>> parameters = new List<KeyValuePair<string, string>>()
                        {
                            new KeyValuePair<string, string>("Identity", identity),
                            new KeyValuePair<string, string>("HiddenFromAddressListsEnabled", hide.Value.ToString())
                        };
                    string sParams = Helper.KeyValuePairListToString(parameters);
                    Trace.AddLog(EventLevel.Information, DateTime.Now, "Exchange", "HideMailContact", "set-mailcontact", "Mail Contact Hidden/Shown", string.Empty, sParams);
                }
                catch (System.Exception ex)
                {
                    string message = ex.Message;
                    string command = ex.InnerException.Data.Contains("CommandName") ? ex.InnerException.Data["CommandName"].ToString() : string.Empty;
                    string function = ex.InnerException.Data.Contains("MethodName") ? ex.InnerException.Data["MethodName"].ToString() : string.Empty;

                    List<KeyValuePair<string, string>> parameters = ex.InnerException.Data.Contains("Parameters") ? (List<KeyValuePair<string, string>>)ex.InnerException.Data["Parameters"] : new List<KeyValuePair<string, string>>();
                    string parameterString = Helper.KeyValuePairListToString(parameters);
                    Trace.AddLog(EventLevel.Error, DateTime.Now, "Exchange", function, command, message, string.Empty, parameterString);
                }


            }

        }

        private void CloseRunspaces()
        {
            
        }

    }
}
