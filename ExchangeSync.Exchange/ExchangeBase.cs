﻿using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Management.Automation;
using System.Management.Automation.Host;
using System.Management.Automation.Runspaces;
using System.Management.Automation.Remoting;
using System.Net;
using System.Text;
using System.Threading.Tasks;

namespace ExchangeSync
{
    public class ExchangeBase : IExchange
    {
        public event ConnectionHandler Connector;
        public delegate void ConnectionHandler(ExchangeBase exch, ExchangeConnectEventArgs e);

        public virtual string AliasPrefix
        {
            get
            {
                throw new NotImplementedException();
            }

            set
            {
                throw new NotImplementedException();
            }
        }

        public virtual string SHELL_URI
        {
            get
            {
                return "http://schemas.microsoft.com/powershell/Microsoft.Exchange";
            }
        }

        public PowerShell psh; public Runspace rs;
        public WSManConnectionInfo connectionInfo;

        public virtual bool AddDistributionListMember(string dlname, string memberidentity)
        {
            bool rc = true;

            Command cmd = new Command("add-distributiongroupmember");
            cmd.Parameters.Add("Identity", dlname);
            cmd.Parameters.Add("Member", memberidentity);

            Pipeline pl = rs.CreatePipeline();
            pl.Commands.Add(cmd);

            try
            {
                Collection<PSObject> results = pl.Invoke();
                if (pl.Error.Count > 0)
                {
                    var error = pl.Error.Read();
                    if (error != null)
                    {
                        List<KeyValuePair<string, string>> parameters = new List<KeyValuePair<string, string>>()
                        {
                            new KeyValuePair<string, string>("Identity", dlname),
                            new KeyValuePair<string, string>("Member", memberidentity)
                        };

                        ExchangeCommandException ex = new ExchangeCommandException();
                        ex.Data.Add("CommandName", "add-distributiongroupmember");
                        ex.Data.Add("MethodName", "AddDistributionListMember");
                        ex.Data.Add("Parameters", parameters);

                        throw new Exception(error.ToString(), ex);
                    }


                    rc = false;
                }
            }
            catch (SystemException ex)
            {
                List<KeyValuePair<string, string>> parameters = new List<KeyValuePair<string, string>>()
                        {
                            new KeyValuePair<string, string>("Identity", dlname),
                            new KeyValuePair<string, string>("Member", memberidentity)
                        };

                ExchangeCommandException exc = new ExchangeCommandException();
                ex.Data.Add("CommandName", "add-distributiongroupmember");
                ex.Data.Add("MethodName", "AddDistributionListMember");
                ex.Data.Add("Parameters", parameters);

                throw new Exception(ex.Message, exc);
            }
            finally
            {
                pl.Error.Close();
                pl.Stop(); pl.Dispose();
            }

            return rc;

        }

        public virtual void ClearPowerShellCommands()
        {
            throw new NotImplementedException();
        }

        public virtual ExchangeConnectEventArgs Connect()
        {
            connectionInfo = GetConnectionInfo(ConfigSetting.ExchangeServerUrl, ConfigSetting.ExchangeCredentials.UserName, ConfigSetting.ExchangeCredentials.Password);
            rs = RunspaceFactory.CreateRunspace(connectionInfo);
            psh = PowerShell.Create();
            psh.Runspace = rs;
            rs.Open();

            bool rc = true;
            RunspaceStateInfo stateInfo = rs.RunspaceStateInfo;
            ExchangeConnectEventArgs e = new ExchangeConnectEventArgs();
            if (stateInfo.State != RunspaceState.Opened)
            {
                e.ConnectionState = stateInfo.State;
                e.ConnectionReason = stateInfo.Reason.Message;
                e.Result = false;
            }
            else
            {
                e.ConnectionState = stateInfo.State;
                e.Result = true;
            }
            return e;
        }

        public virtual ExchangeConnectEventArgs Connect(string exchangeServerUrl, NetworkCredential credentials)
        {
            connectionInfo = GetConnectionInfo(exchangeServerUrl, credentials.UserName, credentials.Password);
            rs = RunspaceFactory.CreateRunspace(connectionInfo);
            psh = PowerShell.Create();
            psh.Runspace = rs;
            rs.Open();

            bool rc = true;
            RunspaceStateInfo stateInfo = rs.RunspaceStateInfo;
            ExchangeConnectEventArgs e = new ExchangeConnectEventArgs();
            if (stateInfo.State != RunspaceState.Opened)
            {
                e.ConnectionState = stateInfo.State;
                e.ConnectionReason = stateInfo.Reason.Message;
                e.Result = false;
            }
            else
            {
                e.ConnectionState = stateInfo.State;
                e.Result = true;
            }
            return e;
        }

        private WSManConnectionInfo GetConnectionInfo(string servername, string username, string password) //, out string[] servers)
        {
            System.Uri serveruri = new Uri(String.Format("http://{0}/powershell?serializationLevel=Full", servername));
            PSCredential creds;

            if (username.Length > 0 && password.Length > 0)
            {
                System.Security.SecureString securePassword = new System.Security.SecureString();

                foreach (char c in password.ToCharArray())
                {
                    securePassword.AppendChar(c);
                }
                creds = new PSCredential(username, securePassword);
            }
            else
            {
                // Use Windows Authentication
                creds = (PSCredential)null;
            }

            WSManConnectionInfo connectionInfo = new WSManConnectionInfo(serveruri, SHELL_URI, creds);
            connectionInfo.OperationTimeout = 5 * 60 * 1000; // 5 minutes.
            connectionInfo.OpenTimeout = 2 * 60 * 1000; // 2 minute.
            connectionInfo.IdleTimeout = 3 * 60 * 1000; // 3 minute.
            connectionInfo.CancelTimeout = 1 * 60 * 1000; // 1 minute.
            return connectionInfo;
        }

        public virtual bool ContactExists(string emailAddressOrAlias)
        {
            bool rc = false;

            Command cmd = new Command("get-mailcontact");
            cmd.Parameters.Add("Identity", emailAddressOrAlias);

            Pipeline pl = rs.CreatePipeline();
            pl.Commands.Add(cmd);

            try
            {
                Collection<PSObject> results = pl.Invoke();
                if (pl.Error.Count > 0)
                {
                    var error = pl.Error.Read();
                    if (error != null)
                    {
                        List<KeyValuePair<string, string>> parameters = new List<KeyValuePair<string, string>>()
                        {
                            new KeyValuePair<string, string>("Identity", emailAddressOrAlias)
                        };

                        ExchangeCommandException ex = new ExchangeCommandException();
                        ex.Data.Add("CommandName", "get-mailcontact");
                        ex.Data.Add("MethodName", "ContactExists");
                        ex.Data.Add("Parameters", parameters);

                        throw new Exception(error.ToString(), ex);
                    }

                    rc = false;
                }

                if (results.Count > 0)
                {
                    rc = true;
                    if (results.Count > 1)
                    {
                        Trace.AddLog(EventLevel.Warning, DateTime.Now, "ContactExists", "Exchange", "Multiple Results for Email Address/Alias " + emailAddressOrAlias, string.Empty);
                    }
                }

            }
            catch (SystemException ex)
            {
                List<KeyValuePair<string, string>> parameters = new List<KeyValuePair<string, string>>()
                {
                    new KeyValuePair<string, string>("Identity", emailAddressOrAlias)
                };

                ExchangeCommandException exc = new ExchangeCommandException();
                ex.Data.Add("CommandName", "get-mailcontact");
                ex.Data.Add("MethodName", "ContactExists");
                ex.Data.Add("Parameters", parameters);

                throw new Exception(ex.Message, exc);
            }
            finally
            {
                pl.Error.Close();
                pl.Stop(); pl.Dispose();
            }

            return rc;

        }

        public virtual bool CreateDistributionGroup(string dlname, string ou)
        {
            return CreateDistributionGroup(dlname, ou, GroupType.Distribution);
        }

        public virtual bool CreateDistributionGroup(string dlname, string ou, GroupType groupType)
        {
            bool rc = true;

            Command cmd = new Command("new-DistributionGroup");
            cmd.Parameters.Add("Name", dlname);
            cmd.Parameters.Add("DisplayName", dlname);
            string alias = ExchangeHelper.ValidateAlias(dlname);
            cmd.Parameters.Add("Alias", alias);
            cmd.Parameters.Add("Type", groupType.ToString());

            if (!string.IsNullOrEmpty(ou))
                cmd.Parameters.Add("OrganizationalUnit", ou);

            Pipeline pl = rs.CreatePipeline();
            pl.Commands.Add(cmd);

            try
            {
                Collection<PSObject> results = pl.Invoke();
                if (pl.Error.Count > 0)
                {
                    var error = pl.Error.Read();
                    if (error != null)
                    {
                        List<KeyValuePair<string, string>> parameters = new List<KeyValuePair<string, string>>()
                        {
                            new KeyValuePair<string, string>("Name", dlname),
                            new KeyValuePair<string, string>("DisplayName", dlname),
                            new KeyValuePair<string, string>("Alias", alias),
                            new KeyValuePair<string, string>("Type", "Distribution")

                        };

                        ExchangeCommandException ex = new ExchangeCommandException();
                        ex.Data.Add("CommandName", "new-DistributionGroup");
                        ex.Data.Add("MethodName", "CreateDistributionGroup");
                        ex.Data.Add("Parameters", parameters);

                        throw new Exception(error.ToString(), ex);
                    }

                    rc = false;
                }
            }
            catch (SystemException ex)
            {
                List<KeyValuePair<string, string>> parameters = new List<KeyValuePair<string, string>>()
                {
                    new KeyValuePair<string, string>("Name", dlname),
                    new KeyValuePair<string, string>("DisplayName", dlname),
                    new KeyValuePair<string, string>("Alias", dlname),
                    new KeyValuePair<string, string>("Type", "Distribution")

                };

                ExchangeCommandException exc = new ExchangeCommandException();
                ex.Data.Add("CommandName", "new-DistributionGroup");
                ex.Data.Add("MethodName", "CreateDistributionGroup");
                ex.Data.Add("Parameters", parameters);

                throw new Exception(ex.Message, exc);

            }
            finally
            {
                pl.Error.Close();
                pl.Stop(); pl.Dispose();
            }

            return rc;

        }

        public virtual bool CreateMailContact(string emailaddress, string ou)
        {
            bool rc = true;
            string name = ExchangeHelper.GenerateNameFromEmailAddress(emailaddress);

            Command cmd = new Command("new-mailcontact");
            cmd.Parameters.Add("Name", name);
            cmd.Parameters.Add("DisplayName", emailaddress);
            cmd.Parameters.Add("ExternalEmailAddress", emailaddress);
            cmd.Parameters.Add("PrimarySmtpAddress", emailaddress);
            cmd.Parameters.Add("FirstName", "");
            cmd.Parameters.Add("LastName", "");
            cmd.Parameters.Add("Alias", name);

            Pipeline pl = rs.CreatePipeline();
            pl.Commands.Add(cmd);

            try
            {
                Collection<PSObject> results = pl.Invoke();
                if (pl.Error.Count > 0)
                {
                    var error = pl.Error.Read();
                    if (error != null)
                    {
                        List<KeyValuePair<string, string>> parameters = new List<KeyValuePair<string, string>>()
                        {
                            new KeyValuePair<string, string>("Name", name),
                            new KeyValuePair<string, string>("DisplayName", emailaddress),
                            new KeyValuePair<string, string>("ExternalEmailAddress", emailaddress),
                            new KeyValuePair<string, string>("FirstName", string.Empty),
                            new KeyValuePair<string, string>("LastName", string.Empty),
                            new KeyValuePair<string, string>("Alias", name)
                        };

                        ExchangeCommandException ex = new ExchangeCommandException();
                        ex.Data.Add("CommandName", "new-mailcontact");
                        ex.Data.Add("MethodName", "CreateMailContact");
                        ex.Data.Add("Parameters", parameters);

                        throw new Exception(error.ToString(), ex);
                    }

                    rc = false;
                }
            }
            catch (SystemException ex)
            {
                rc = false;

                List<KeyValuePair<string, string>> parameters = new List<KeyValuePair<string, string>>()
                {
                    new KeyValuePair<string, string>("Name", name),
                    new KeyValuePair<string, string>("DisplayName", emailaddress),
                    new KeyValuePair<string, string>("ExternalEmailAddress", emailaddress),
                    new KeyValuePair<string, string>("FirstName", string.Empty),
                    new KeyValuePair<string, string>("LastName", string.Empty),
                    new KeyValuePair<string, string>("Alias", name)
                };

                ExchangeCommandException exc = new ExchangeCommandException();
                ex.Data.Add("CommandName", "new-mailcontact");
                ex.Data.Add("MethodName", "CreateMailContact");
                ex.Data.Add("Parameters", parameters);

                throw new Exception(ex.Message, exc);

            }
            finally
            {
                pl.Error.Close();
                pl.Stop(); pl.Dispose();
            }

            return rc;

        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="displayName"></param>
        /// <param name="externalEmailAddress">Real Email Address to be used</param>
        /// <param name="primarySmtpAddress">Fake Email Address to be used</param>
        /// <param name="firstName"></param>
        /// <param name="lastName"></param>
        /// <param name="ou"></param>
        /// <param name="alias"></param>
        /// <returns></returns>
        public virtual bool CreateMailContact(string displayName, string externalEmailAddress, string firstName, string lastName, string ou, out string alias, string primarySmtpAddress = "")
        {
            bool rc = true;
            if (!string.IsNullOrEmpty(displayName))
                displayName = ExchangeHelper.TrimLength(displayName, 64);
            else
                displayName = externalEmailAddress;

            string fullname = ExchangeHelper.TrimLength(displayName, 64);
            string name = ExchangeHelper.FormatName(firstName, lastName, externalEmailAddress, string.Empty);
            alias = ExchangeHelper.FormatAlias(firstName, lastName, externalEmailAddress, string.Empty);

            Command cmd = new Command("new-mailcontact");
            cmd.Parameters.Add("Name", name);
            // cmd.Parameters.Add("DisplayName", displayname);
            cmd.Parameters.Add("DisplayName", fullname);
            cmd.Parameters.Add("ExternalEmailAddress", externalEmailAddress);
            cmd.Parameters.Add("PrimarySmtpAddress", primarySmtpAddress == "" ? externalEmailAddress : primarySmtpAddress);
            cmd.Parameters.Add("FirstName", firstName);
            cmd.Parameters.Add("LastName", lastName);
            cmd.Parameters.Add("Alias", alias);

            if (!string.IsNullOrEmpty(ou))
                cmd.Parameters.Add("OrganizationalUnit", ou);
                
            Pipeline pl = rs.CreatePipeline();
            pl.Commands.Add(cmd);

            try
            {
                Collection<PSObject> results = pl.Invoke();
                if (pl.Error.Count > 0)
                {
                    var error = pl.Error.Read();
                    if (error != null)
                    {
                        List<KeyValuePair<string, string>> parameters = new List<KeyValuePair<string, string>>()
                        {
                            new KeyValuePair<string, string>("Name", name),
                            new KeyValuePair<string, string>("DisplayName", fullname),
                            new KeyValuePair<string, string>("ExternalEmailAddress", externalEmailAddress),
                            new KeyValuePair<string, string>("FirstName", firstName),
                            new KeyValuePair<string, string>("LastName", lastName),
                            new KeyValuePair<string, string>("Alias", alias),
                            new KeyValuePair<string, string>("OrganizationalUnit", ou)
                        };

                        ExchangeCommandException ex = new ExchangeCommandException();
                        ex.Data.Add("CommandName", "new-mailcontact");
                        ex.Data.Add("MethodName", "CreateMailContact");
                        ex.Data.Add("Parameters", parameters);

                        throw new Exception(error.ToString(), ex);
                    }

                    rc = false;
                }
            }
            catch (SystemException ex)
            {
                rc = false;
                List<KeyValuePair<string, string>> parameters = new List<KeyValuePair<string, string>>()
                {
                    new KeyValuePair<string, string>("Name", name),
                    new KeyValuePair<string, string>("DisplayName", fullname),
                    new KeyValuePair<string, string>("ExternalEmailAddress", externalEmailAddress),
                    new KeyValuePair<string, string>("FirstName", firstName),
                    new KeyValuePair<string, string>("LastName", lastName),
                    new KeyValuePair<string, string>("Alias", alias),
                    new KeyValuePair<string, string>("OrganizationalUnit", ou)
                };

                ExchangeCommandException exc = new ExchangeCommandException();
                ex.Data.Add("CommandName", "new-mailcontact");
                ex.Data.Add("MethodName", "CreateMailContact");
                ex.Data.Add("Parameters", parameters);

                throw new Exception(ex.Message, exc);

            }
            finally
            {
                pl.Error.Close();
                pl.Stop(); pl.Dispose();
            }
            return rc;

        }

        public virtual bool CreateMailContact(string displayName, string externalEmailAddress, string firstName, string lastName, string ou, string custom, out string alias, string primarySmtpAddress = "")
        {
            bool rc = true;
            if (!string.IsNullOrEmpty(displayName))
                displayName = ExchangeHelper.TrimLength(displayName, 64);
            else
                displayName = externalEmailAddress;

            string name = ExchangeHelper.FormatName(firstName, lastName, externalEmailAddress, custom);
            alias = ExchangeHelper.FormatAlias(firstName, lastName, externalEmailAddress, custom);

            Command cmd = new Command("new-mailcontact");
            cmd.Parameters.Add("Name", name);
            cmd.Parameters.Add("DisplayName", displayName);
            cmd.Parameters.Add("ExternalEmailAddress", externalEmailAddress);
            cmd.Parameters.Add("PrimarySmtpAddress", primarySmtpAddress == "" ? externalEmailAddress : primarySmtpAddress);
            cmd.Parameters.Add("FirstName", firstName);
            cmd.Parameters.Add("LastName", lastName);
            cmd.Parameters.Add("Alias", alias);
            
            if (!string.IsNullOrEmpty(ou))
                cmd.Parameters.Add("OrganizationalUnit", ou);

            Pipeline pl = rs.CreatePipeline();
            pl.Commands.Add(cmd);

            try
            {
                Collection<PSObject> results = pl.Invoke();
                if (pl.Error.Count > 0)
                {
                    var  error = pl.Error.Read();
                    if (error != null)
                    {
                        List<KeyValuePair<string, string>> parameters = new List<KeyValuePair<string, string>>()
                        {
                            new KeyValuePair<string, string>("Name", name),
                            new KeyValuePair<string, string>("DisplayName", displayName),
                            new KeyValuePair<string, string>("ExternalEmailAddress", externalEmailAddress),
                            new KeyValuePair<string, string>("FirstName", firstName),
                            new KeyValuePair<string, string>("LastName", lastName),
                            new KeyValuePair<string, string>("Alias", alias),
                            new KeyValuePair<string, string>("OrganizationalUnit", ou)
                        };

                        ExchangeCommandException ex = new ExchangeCommandException();
                        ex.Data.Add("CommandName", "new-mailcontact");
                        ex.Data.Add("MethodName", "CreateMailContact");
                        ex.Data.Add("Parameters", parameters);

                        throw new Exception(error.ToString(), ex);
                    }
                    
                    rc = false;
                    
                }
            }
            catch (SystemException ex)
            {
                rc = false;
                List<KeyValuePair<string, string>> parameters = new List<KeyValuePair<string, string>>()
                {
                    new KeyValuePair<string, string>("Name", name),
                    new KeyValuePair<string, string>("DisplayName", displayName),
                    new KeyValuePair<string, string>("ExternalEmailAddress", externalEmailAddress),
                    new KeyValuePair<string, string>("FirstName", firstName),
                    new KeyValuePair<string, string>("LastName", lastName),
                    new KeyValuePair<string, string>("Alias", alias),
                    new KeyValuePair<string, string>("OrganizationalUnit", ou)
                };

                ExchangeCommandException exc = new ExchangeCommandException();
                ex.Data.Add("CommandName", "new-mailcontact");
                ex.Data.Add("MethodName", "CreateMailContact");
                ex.Data.Add("Parameters", parameters);

                throw new Exception(ex.Message, exc);

            }
            finally
            {
                pl.Error.Close();
                pl.Stop(); pl.Dispose();
            }
            return rc;

        }

        public virtual bool Disconnect()
        {
            rs.Close();
            rs.Dispose();
            return true;
        }

        public virtual bool DistributionGroupExists(string dlname)
        {
            bool rc = false;
            Command cmd = new Command("get-distributiongroup");
            cmd.Parameters.Add("Identity", dlname);

            Pipeline pl = rs.CreatePipeline();
            pl.Commands.Add(cmd);

            try
            {
                Collection<PSObject> results = pl.Invoke();
                if (pl.Error.Count > 0)
                {
                    var error = pl.Error.Read();
                    if (error != null)
                    {
                        List<KeyValuePair<string, string>> parameters = new List<KeyValuePair<string, string>>()
                        {
                            new KeyValuePair<string, string>("Identity", dlname)
                        };

                        ExchangeCommandException ex = new ExchangeCommandException();
                        ex.Data.Add("CommandName", "get-distributiongroup");
                        ex.Data.Add("MethodName", "DistributionGroupExists");
                        ex.Data.Add("Parameters", parameters);

                        throw new Exception(error.ToString(), ex);
                    }

                    rc = false;
                }

                if (results.Count > 0)
                {
                    rc = true;
                }
            }
            catch (SystemException ex)
            {
                rc = false;
                List<KeyValuePair<string, string>> parameters = new List<KeyValuePair<string, string>>()
                {
                    new KeyValuePair<string, string>("Identity", dlname)
                };

                ExchangeCommandException exc = new ExchangeCommandException();
                ex.Data.Add("CommandName", "get-distributiongroup");
                ex.Data.Add("MethodName", "DistributionGroupExists");
                ex.Data.Add("Parameters", parameters);

                throw new Exception(ex.Message, exc);

            }
            finally
            {
                pl.Error.Close();
                pl.Stop(); pl.Dispose();
            }

            return rc;

        }

        public virtual bool ExchangeUserExists(string identity, bool mailboxEnabled)
        {
            bool rc = false;

            string commandName = "";
            if (mailboxEnabled)
                commandName = "get-mailbox";
            else
                commandName = "get-user";

            Command cmd = new Command(commandName);
            cmd.Parameters.Add("Identity", identity);

            Pipeline pl = rs.CreatePipeline();
            pl.Commands.Add(cmd);

            try
            {
                Collection<PSObject> results = pl.Invoke();
                if (pl.Error.Count > 0)
                {
                    var error = pl.Error.Read();
                    if (error != null)
                    {
                        List<KeyValuePair<string, string>> parameters = new List<KeyValuePair<string, string>>()
                        {
                            new KeyValuePair<string, string>("Identity", identity)
                        };

                        ExchangeCommandException ex = new ExchangeCommandException();
                        ex.Data.Add("CommandName", commandName);
                        ex.Data.Add("MethodName", "ExchangeUserExists");
                        ex.Data.Add("Parameters", parameters);

                        throw new Exception(error.ToString(), ex);
                    }

                    rc = false;
                }

                if (results.Count > 0)
                {
                    rc = true;
                    if (results.Count > 1)
                    {
                        Trace.AddLog(EventLevel.Warning, DateTime.Now, "ExchangeUserExists", "Exchange", "Multiple Results for User with Identity " + identity, string.Empty);
                    }
                }

            }
            catch (SystemException ex)
            {
                List<KeyValuePair<string, string>> parameters = new List<KeyValuePair<string, string>>()
                {
                    new KeyValuePair<string, string>("Identity", identity)
                };

                ExchangeCommandException exc = new ExchangeCommandException();
                ex.Data.Add("CommandName", commandName);
                ex.Data.Add("MethodName", "ExchangeUserExists");
                ex.Data.Add("Parameters", parameters);

                throw new Exception(ex.Message, exc);
            }
            finally
            {
                pl.Error.Close();
                pl.Stop(); pl.Dispose();
            }

            return rc;

        }

        public virtual List<DistributionGroup> GetAllDistributionListGroups()
        {
            List<DistributionGroup> rc = new List<DistributionGroup>();
            Command cmd = new Command("get-distributiongroup");

            Pipeline pl = rs.CreatePipeline();
            pl.Commands.Add(cmd);

            try
            {
                Collection<PSObject> results = pl.Invoke();
                if (pl.Error.Count > 0)
                {
                    var error = pl.Error.Read();
                    if (error != null)
                    {
                        List<KeyValuePair<string, string>> parameters = new List<KeyValuePair<string, string>>();
                        
                        ExchangeCommandException ex = new ExchangeCommandException();
                        ex.Data.Add("CommandName", "get-distributiongroup");
                        ex.Data.Add("MethodName", "GetAllDistributionListGroups");
                        ex.Data.Add("Parameters", parameters);

                        throw new Exception(error.ToString(), ex);
                    }
                }

                if (results.Count > 0)
                {
                    foreach (PSObject result in results)
                    {
                        string name = "", ou = "";
                        if (result.Members["Name"].Value != null)
                            name = result.Members["Name"].Value.ToString();

                        // if (result.Members["OrganizationalUnit"].Value != null)
                        //    ou = result.Members["OrganizationalUnit"].Value.ToString();

                        DistributionGroup group = new DistributionGroup();
                        group.Add(name, string.Empty);
                        rc.Add(group);
                    }
                }
            }
            catch (SystemException ex)
            {
                List<KeyValuePair<string, string>> parameters = new List<KeyValuePair<string, string>>();
                ExchangeCommandException exc = new ExchangeCommandException();
                ex.Data.Add("CommandName", "get-distributiongroup");
                ex.Data.Add("MethodName", "GetAllDistributionListGroups");
                ex.Data.Add("Parameters", parameters);

                throw new Exception(ex.Message, exc);
            }
            finally
            {
                pl.Error.Close();
                pl.Stop(); pl.Dispose();
            }

            return rc;

        }

        public virtual List<string> GetAllDistributionLists()
        {
            List<string> rc = new List<string>();
            Command cmd = new Command("get-distributiongroup");

            Pipeline pl = rs.CreatePipeline();
            pl.Commands.Add(cmd);

            try
            {
                Collection<PSObject> results = pl.Invoke();
                if (pl.Error.Count > 0)
                {
                    var error = pl.Error.Read();
                    if (error != null)
                    {
                        List<KeyValuePair<string, string>> parameters = new List<KeyValuePair<string, string>>();

                        ExchangeCommandException ex = new ExchangeCommandException();
                        ex.Data.Add("CommandName", "get-distributiongroup");
                        ex.Data.Add("MethodName", "GetAllDistributionLists");
                        ex.Data.Add("Parameters", parameters);

                        throw new Exception(error.ToString(), ex);
                    }
                }

                if (results.Count > 0)
                {
                    foreach (PSObject result in results)
                    {
                        string name = "", ou = "";
                        if (result.Members["Name"].Value != null)
                        {
                            rc.Add(result.Members["Name"].Value.ToString());
                        }

                    }
                }
            }
            catch (SystemException ex)
            {
                List<KeyValuePair<string, string>> parameters = new List<KeyValuePair<string, string>>();
                ExchangeCommandException exc = new ExchangeCommandException();
                ex.Data.Add("CommandName", "get-distributiongroup");
                ex.Data.Add("MethodName", "GetAllDistributionListGroups");
                ex.Data.Add("Parameters", parameters);

                throw new Exception(ex.Message, exc);

            }
            finally
            {
                pl.Error.Close();
                pl.Stop(); pl.Dispose();
            }

            return rc;

        }

        public virtual List<OrganizationalUnit> GetAllOrganizationalUnits()
        {
            List<OrganizationalUnit> rc = new List<OrganizationalUnit>();

            Command cmd = new Command("get-organizationalunit");

            Pipeline pl = rs.CreatePipeline();
            pl.Commands.Add(cmd);

            try
            {
                Collection<PSObject> results = pl.Invoke();
                if (pl.Error.Count > 0)
                {
                    var error = pl.Error.Read();
                    if (error != null)
                    {
                        List<KeyValuePair<string, string>> parameters = new List<KeyValuePair<string, string>>();

                        ExchangeCommandException ex = new ExchangeCommandException();
                        ex.Data.Add("CommandName", "get-organizationalunit");
                        ex.Data.Add("MethodName", "GetAllOrganizationalUnits");
                        ex.Data.Add("Parameters", parameters);

                        throw new Exception(error.ToString(), ex);
                    }
                }

                if (results.Count > 0)
                {
                    foreach (PSObject result in results)
                    {
                        string name = "", canonicalname = "";
                        if (result.Members["Name"].Value != null)
                            name = result.Members["Name"].Value.ToString();

                        if (result.Members["CanonicalName"].Value != null)
                            canonicalname = result.Members["CanonicalName"].Value.ToString();

                        OrganizationalUnit unit = new OrganizationalUnit();
                        unit.Add(name, canonicalname);
                        rc.Add(unit);

                    }
                }
            }
            catch (SystemException ex)
            {
                List<KeyValuePair<string, string>> parameters = new List<KeyValuePair<string, string>>();
                ExchangeCommandException exc = new ExchangeCommandException();
                ex.Data.Add("CommandName", "get-organizationalunit");
                ex.Data.Add("MethodName", "GetAllOrganizationalUnits");
                ex.Data.Add("Parameters", parameters);

                throw new Exception(ex.Message, exc);
            }
            finally
            {
                pl.Error.Close();
                pl.Stop(); pl.Dispose();
            }

            return rc;
        }

        public virtual List<ExchangeContact> GetDistributionListMembers(string dlname)
        {
            List<ExchangeContact> rc = new List<ExchangeContact>();
            Command cmd = new Command("get-distributiongroupmember");
            cmd.Parameters.Add("Identity", dlname);

            Pipeline pl = rs.CreatePipeline();
            pl.Commands.Add(cmd);

            try
            {
                Collection<PSObject> results = pl.Invoke();
                if (pl.Error.Count > 0)
                {
                    var error = pl.Error.Read();
                    if (error != null)
                    {
                        List<KeyValuePair<string, string>> parameters = new List<KeyValuePair<string, string>>()
                        {
                            new KeyValuePair<string, string>("Identity", dlname)
                        };

                        ExchangeCommandException ex = new ExchangeCommandException();
                        ex.Data.Add("CommandName", "get-distributiongroupmember");
                        ex.Data.Add("MethodName", "GetDistributionListMembers");
                        ex.Data.Add("Parameters", parameters);

                        throw new Exception(error.ToString(), ex);
                    }
                }

                if (results.Count > 0)
                {
                    foreach (PSObject result in results)
                    {
                        ExchangeContact contact = new ExchangeContact();
                        if (result.Members["Name"].Value != null)
                            contact.Name = result.Members["Name"].Value.ToString();

                        if (result.Members["Alias"].Value != null)
                            contact.Alias = result.Members["Alias"].Value.ToString();

                        if (result.Members["OrganizationalUnit"].Value != null)
                            contact.OrganizationalUnit = result.Members["OrganizationalUnit"].Value.ToString();

                        if (result.Members["PrimarySmtpAddress"].Value != null)
                            contact.PrimarySmtpAddress = result.Members["PrimarySmtpAddress"].Value.ToString();

                        if (result.Members["Identity"].Value != null)
                            contact.Identity = result.Members["Identity"].Value.ToString();

                        rc.Add(contact);
                    }
                }
            }
            catch (SystemException ex)
            {
                List<KeyValuePair<string, string>> parameters = new List<KeyValuePair<string, string>>()
                {
                    new KeyValuePair<string, string>("Identity", dlname)
                };

                ExchangeCommandException exc = new ExchangeCommandException();
                ex.Data.Add("CommandName", "get-distributiongroupmember");
                ex.Data.Add("MethodName", "GetDistributionListMembers");
                ex.Data.Add("Parameters", parameters);

                throw new Exception(ex.Message, exc);

            }
            finally
            {
                pl.Error.Close();
                pl.Stop(); pl.Dispose();
            }

            return rc;

        }

        public virtual ExchangeContact GetRecipient(string identity)
        {
            ExchangeContact contact = new ExchangeContact();

            Command cmd = new Command("get-recipient");
            cmd.Parameters.Add("Identity", identity);

            Pipeline pl = rs.CreatePipeline();
            pl.Commands.Add(cmd);

            try
            {
                Collection<PSObject> results = pl.Invoke();
                if (pl.Error.Count > 0)
                {
                    var error = pl.Error.Read();
                    if (error != null)
                    {
                        List<KeyValuePair<string, string>> parameters = new List<KeyValuePair<string, string>>()
                        {
                            new KeyValuePair<string, string>("Identity", identity)
                        };

                        ExchangeCommandException ex = new ExchangeCommandException();
                        ex.Data.Add("CommandName", "get-recipient");
                        ex.Data.Add("MethodName", "GetRecipient");
                        ex.Data.Add("Parameters", parameters);

                        throw new Exception(error.ToString(), ex);
                    }
                }

                if (results.Count > 0)
                {
                    foreach (PSObject result in results)
                    {
                        if (result.Members["Name"].Value != null)
                            contact.Name = result.Members["Name"].Value.ToString();

                        if (result.Members["Alias"].Value != null)
                            contact.Alias = result.Members["Alias"].Value.ToString();

                        if (result.Members["OrganizationalUnit"].Value != null)
                            contact.OrganizationalUnit = result.Members["OrganizationalUnit"].Value.ToString();

                        if (result.Members["PrimarySmtpAddress"].Value != null)
                            contact.PrimarySmtpAddress = result.Members["PrimarySmtpAddress"].Value.ToString();

                        if (result.Members["Identity"].Value != null)
                            contact.Identity = result.Members["Identity"].Value.ToString();

                        break;
                    }
                }
            }
            catch (SystemException ex)
            {
                List<KeyValuePair<string, string>> parameters = new List<KeyValuePair<string, string>>()
                {
                    new KeyValuePair<string, string>("Identity", identity)
                };

                ExchangeCommandException exc = new ExchangeCommandException();
                ex.Data.Add("CommandName", "get-mailcontact");
                ex.Data.Add("MethodName", "GetMailContact");
                ex.Data.Add("Parameters", parameters);

                throw new Exception(ex.Message, exc);

            }
            finally
            {
                pl.Error.Close();
                pl.Stop(); pl.Dispose();
            }

            return contact;

        }

        public virtual ExchangeContact GetMailContact(string identity)
        {
            ExchangeContact contact = new ExchangeContact();

            Command cmd = new Command("get-mailcontact");
            cmd.Parameters.Add("Identity", identity);

            Pipeline pl = rs.CreatePipeline();
            pl.Commands.Add(cmd);

            try
            {
                Collection<PSObject> results = pl.Invoke();
                if (pl.Error.Count > 0)
                {
                    var error = pl.Error.Read();
                    if (error != null)
                    {
                        List<KeyValuePair<string, string>> parameters = new List<KeyValuePair<string, string>>()
                        {
                            new KeyValuePair<string, string>("Identity", identity)
                        };

                        ExchangeCommandException ex = new ExchangeCommandException();
                        ex.Data.Add("CommandName", "get-mailcontact");
                        ex.Data.Add("MethodName", "GetMailContact");
                        ex.Data.Add("Parameters", parameters);

                        throw new Exception(error.ToString(), ex);
                    }
                }

                if (results.Count > 0)
                {
                    foreach (PSObject result in results)
                    {
                        if (result.Members["Name"].Value != null)
                            contact.Name = result.Members["Name"].Value.ToString();

                        if (result.Members["Alias"].Value != null)
                            contact.Alias = result.Members["Alias"].Value.ToString();

                        if (result.Members["OrganizationalUnit"].Value != null)
                            contact.OrganizationalUnit = result.Members["OrganizationalUnit"].Value.ToString();

                        if (result.Members["PrimarySmtpAddress"].Value != null)
                            contact.PrimarySmtpAddress = result.Members["PrimarySmtpAddress"].Value.ToString();

                        if (result.Members["Identity"].Value != null)
                            contact.Identity = result.Members["Identity"].Value.ToString();

                        break;
                    }
                }
            }
            catch (SystemException ex)
            {
                List<KeyValuePair<string, string>> parameters = new List<KeyValuePair<string, string>>()
                {
                    new KeyValuePair<string, string>("Identity", identity)
                };

                ExchangeCommandException exc = new ExchangeCommandException();
                ex.Data.Add("CommandName", "get-mailcontact");
                ex.Data.Add("MethodName", "GetMailContact");
                ex.Data.Add("Parameters", parameters);

                throw new Exception(ex.Message, exc);

            }
            finally
            {
                pl.Error.Close();
                pl.Stop(); pl.Dispose();
            }

            return contact;

        }

        public virtual ExchangeContact GetAllMailContacts(string organizationalUnit = "")
        {
            ExchangeContact contact = new ExchangeContact();

            Command cmd = new Command("get-mailcontact");
            if (!string.IsNullOrEmpty(organizationalUnit))
                cmd.Parameters.Add("OrganizationalUnit", organizationalUnit);

            Pipeline pl = rs.CreatePipeline();
            pl.Commands.Add(cmd);

            try
            {
                Collection<PSObject> results = pl.Invoke();
                if (pl.Error.Count > 0)
                {
                    var error = pl.Error.Read();
                    if (error != null)
                    {
                        List<KeyValuePair<string, string>> parameters = new List<KeyValuePair<string, string>>()
                        {
                            new KeyValuePair<string, string>("OrganizationalUnit", organizationalUnit)
                        };

                        ExchangeCommandException ex = new ExchangeCommandException();
                        ex.Data.Add("CommandName", "get-mailcontact");
                        ex.Data.Add("MethodName", "GetMailContact");
                        ex.Data.Add("Parameters", parameters);

                        throw new Exception(error.ToString(), ex);
                    }
                }

                if (results.Count > 0)
                {
                    foreach (PSObject result in results)
                    {
                        if (result.Members["Name"].Value != null)
                            contact.Name = result.Members["Name"].Value.ToString();

                        if (result.Members["Alias"].Value != null)
                            contact.Alias = result.Members["Alias"].Value.ToString();

                        if (result.Members["OrganizationalUnit"].Value != null)
                            contact.OrganizationalUnit = result.Members["OrganizationalUnit"].Value.ToString();

                        if (result.Members["PrimarySmtpAddress"].Value != null)
                            contact.PrimarySmtpAddress = result.Members["PrimarySmtpAddress"].Value.ToString();

                        if (result.Members["Identity"].Value != null)
                            contact.Identity = result.Members["Identity"].Value.ToString();

                        break;
                    }
                }
            }
            catch (SystemException ex)
            {
                List<KeyValuePair<string, string>> parameters = new List<KeyValuePair<string, string>>()
                {
                    new KeyValuePair<string, string>("OrganizationalUnit", organizationalUnit)
                };

                ExchangeCommandException exc = new ExchangeCommandException();
                ex.Data.Add("CommandName", "get-mailcontact");
                ex.Data.Add("MethodName", "GetMailContact");
                ex.Data.Add("Parameters", parameters);

                throw new Exception(ex.Message, exc);

            }
            finally
            {
                pl.Error.Close();
                pl.Stop(); pl.Dispose();
            }

            return contact;

        }



        public virtual ExchangeContact GetMailUser(string identity)
        {
            ExchangeContact contact = new ExchangeContact();
            Command cmd = new Command("get-user");
            cmd.Parameters.Add("Identity", identity);

            Pipeline pl = rs.CreatePipeline();
            pl.Commands.Add(cmd);

            try
            {
                Collection<PSObject> results = pl.Invoke();
                if (pl.Error.Count > 0)
                {
                    var error = pl.Error.Read();
                    if (error != null)
                    {
                        List<KeyValuePair<string, string>> parameters = new List<KeyValuePair<string, string>>()
                        {
                            new KeyValuePair<string, string>("Identity", identity)
                        };

                        ExchangeCommandException ex = new ExchangeCommandException();
                        ex.Data.Add("CommandName", "get-user");
                        ex.Data.Add("MethodName", "GetMailUser");
                        ex.Data.Add("Parameters", parameters);

                        throw new Exception(error.ToString(), ex);
                    }
                }

                if (results.Count > 0)
                {
                    foreach (PSObject result in results)
                    {
                        if (result.Members["Name"].Value != null)
                            contact.Name = result.Members["Name"].Value.ToString();

                        if (result.Members["SamAccountName"].Value != null)
                            contact.Alias = result.Members["SamAccountName"].Value.ToString();

                        // if (result.Members["OrganizationalUnit"].Value != null)
                        contact.OrganizationalUnit = string.Empty; result.Members["OrganizationalUnit"].Value.ToString();

                        if (result.Members["WindowsEmailAddress"].Value != null)
                            contact.PrimarySmtpAddress = result.Members["WindowsEmailAddress"].Value.ToString();

                        if (result.Members["Identity"].Value != null)
                            contact.Identity = result.Members["Identity"].Value.ToString();

                        break;
                    }
                }

            }
            catch (SystemException ex)
            {
                List<KeyValuePair<string, string>> parameters = new List<KeyValuePair<string, string>>()
                {
                    new KeyValuePair<string, string>("Identity", identity)
                };

                ExchangeCommandException exc = new ExchangeCommandException();
                ex.Data.Add("CommandName", "get-user");
                ex.Data.Add("MethodName", "GetMailUser");
                ex.Data.Add("Parameters", parameters);

                throw new Exception(ex.Message, exc);

            }
            finally
            {
                pl.Error.Close();
                pl.Stop(); pl.Dispose();
            }
            return contact;

        }

        public virtual string GetServerName()
        {
            string[] retVal = null;

            Command cmd = new Command("get-exchangeserver");

            Pipeline pl = rs.CreatePipeline();
            pl.Commands.Add(cmd);
            try
            {
                Collection<PSObject> results = pl.Invoke();

                if (pl.Error.Count > 0)
                {
                    var error = pl.Error.Read();
                    if (error != null)
                    {
                        List<KeyValuePair<string, string>> parameters = new List<KeyValuePair<string, string>>()
                        {
                        };

                        ExchangeCommandException ex = new ExchangeCommandException();
                        ex.Data.Add("CommandName", "get-exchangeserver");
                        ex.Data.Add("MethodName", "GetServerName");
                        ex.Data.Add("Parameters", parameters);

                        throw new Exception(error.ToString(), ex);
                    }
                }
                else
                {
                    retVal = new string[1];
                    foreach (PSObject result in results)
                    {
                        if (null == retVal[0] || retVal[0].Length == 0)
                        {
                            retVal[0] = result.Members["Name"].Value as string;
                        }
                        else
                        {
                            string[] newVal = new string[retVal.Length + 1];
                            retVal.CopyTo(newVal, 0);
                            newVal[retVal.Length] = result.Members["Name"].Value as string;
                            retVal = newVal;
                        }
                    }
                }
            }
            catch (SystemException ex)
            {
                List<KeyValuePair<string, string>> parameters = new List<KeyValuePair<string, string>>()
                {
                };

                ExchangeCommandException exc = new ExchangeCommandException();
                ex.Data.Add("CommandName", "get-exchangeserver");
                ex.Data.Add("MethodName", "GetServerName");
                ex.Data.Add("Parameters", parameters);

                throw new Exception(ex.Message, exc);

            }
            finally
            {
                pl.Error.Close();
                pl.Stop(); pl.Dispose();
            }
            return retVal[0].ToString();
        }

        public virtual bool HideMailContact(string identity, bool hidden)
        {
            bool rc = true;

            Command cmd = new Command("set-mailcontact");
            cmd.Parameters.Add("Identity", identity);
            cmd.Parameters.Add("HiddenFromAddressListsEnabled", true);
            cmd.Parameters.Add("ForceUpgrade");

            Pipeline pl = rs.CreatePipeline();
            pl.Commands.Add(cmd);

            try
            {
                Collection<PSObject> results = pl.Invoke();
                if (pl.Error.Count > 0)
                {
                    var error = pl.Error.Read();
                    if (error != null)
                    {
                        List<KeyValuePair<string, string>> parameters = new List<KeyValuePair<string, string>>()
                        {
                            new KeyValuePair<string, string>("Identity", identity),
                            new KeyValuePair<string, string>("HiddenFromAddressListsEnabled", true.ToString())
                        };

                        ExchangeCommandException ex = new ExchangeCommandException();
                        ex.Data.Add("CommandName", "new-mailcontact");
                        ex.Data.Add("MethodName", "AddDistributionListMember");
                        ex.Data.Add("Parameters", parameters);

                        throw new Exception(error.ToString(), ex);
                    }

                    rc = false;
                }
            }
            catch (SystemException ex)
            {
                rc = false;
                List<KeyValuePair<string, string>> parameters = new List<KeyValuePair<string, string>>()
                {
                    new KeyValuePair<string, string>("Identity", identity),
                    new KeyValuePair<string, string>("HiddenFromAddressListsEnabled", true.ToString())
                };

                ExchangeCommandException exc = new ExchangeCommandException();
                ex.Data.Add("CommandName", "new-mailcontact");
                ex.Data.Add("Parameters", parameters);

                throw new Exception(ex.Message, exc);

            }
            finally
            {
                pl.Error.Close();
                pl.Stop(); pl.Dispose();
            }

            return rc;
        }

        public virtual bool IsOUValid(string ouName)
        {
            bool rc = false;

            Command cmd = new Command("get-organizationalunit");
            cmd.Parameters.Add("Identity", ouName);

            Pipeline pl = rs.CreatePipeline();
            pl.Commands.Add(cmd);

            try
            {
                Collection<PSObject> results = pl.Invoke();
                if (pl.Error.Count > 0)
                {
                    var error = pl.Error.Read();
                    if (error != null)
                    {
                        List<KeyValuePair<string, string>> parameters = new List<KeyValuePair<string, string>>()
                        {
                            new KeyValuePair<string, string>("Identity", ouName),
                        };

                        ExchangeCommandException ex = new ExchangeCommandException();
                        ex.Data.Add("CommandName", "get-organizationalunit");
                        ex.Data.Add("MethodName", "IsOUValid");
                        ex.Data.Add("Parameters", parameters);

                        throw new Exception(error.ToString(), ex);
                    }

                    rc = false;
                }

                if (results.Count > 0)
                {
                    foreach (PSObject result in results)
                    {
                        if (result.Members["Name"].Value != null)
                        {
                            string value = result.Members["IsValid"].Value.ToString().ToLower();
                            if (value == "true")
                                rc = true;
                        }
                    }
                }
            }
            catch (SystemException ex)
            {
                rc = false;
                List<KeyValuePair<string, string>> parameters = new List<KeyValuePair<string, string>>()
                {
                    new KeyValuePair<string, string>("Identity", ouName),
                };

                ExchangeCommandException exc = new ExchangeCommandException();
                ex.Data.Add("CommandName", "get-organizationalunit");
                ex.Data.Add("Parameters", parameters);

                throw new Exception(ex.Message, exc);

            }
            finally
            {
                pl.Error.Close();
                pl.Stop(); pl.Dispose();
            }

            return rc;

        }

        public virtual void RemoveAllDistributionListMembers(string dlname)
        {
            List<ExchangeContact> members = GetDistributionListMembers(dlname);
            Trace.AddLog(EventLevel.Information, DateTime.Now, "RemoveAllDistributionListMembers", "Exchange", "Calling Remove All Distribution List Members", members.Count.ToString() + " members retrieved");
            if (members.Count > 0)
            {
                foreach (ExchangeContact contact in members)
                {
                    try
                    {
                        if (!string.IsNullOrEmpty(contact.Alias))
                        {
                            RemoveDistributionListMember(dlname, contact.Alias);
                        }
                        else if (!string.IsNullOrEmpty(contact.PrimarySmtpAddress))
                        {
                            RemoveDistributionListMember(dlname, contact.PrimarySmtpAddress);
                        }
                        else
                        {
                            RemoveDistributionListMember(dlname, contact.Identity);
                        }

                        if (ExchangeSync.AppSetting.LoggingCriteriaVerbose.GetValue().ToLower() == "yes")
                        {
                            List<KeyValuePair<string, string>> parameters = new List<KeyValuePair<string, string>>()
                            {
                                new KeyValuePair<string, string>("Identity", dlname),
                                new KeyValuePair<string, string>("Member", contact.Alias)
                            };
                            string sParams = Helper.KeyValuePairListToString(parameters);
                            Trace.AddLog(EventLevel.Verbose, DateTime.Now, "Exchange", "RemoveDistributionListMember", "remove-distributiongroupmember", "Remove Distribution Group Member", string.Empty, sParams);
                        }
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

        }

        public virtual void RemoveDistributionGroup(string groupName)
        {
            Command cmd = new Command("remove-distributiongroup");
            cmd.Parameters.Add("Identity", groupName);
            cmd.Parameters.Add("Confirm", false);

            Pipeline pl = rs.CreatePipeline();
            pl.Commands.Add(cmd);

            try
            {
                Collection<PSObject> results = pl.Invoke();
                if (pl.Error.Count > 0)
                {
                    var error = pl.Error.Read();
                    if (error != null)
                    {
                        List<KeyValuePair<string, string>> parameters = new List<KeyValuePair<string, string>>()
                        {
                            new KeyValuePair<string, string>("Identity", groupName),
                            new KeyValuePair<string, string>("Confirm", false.ToString())
                        };

                        ExchangeCommandException ex = new ExchangeCommandException();
                        ex.Data.Add("CommandName", "remove-distributiongroup");
                        ex.Data.Add("MethodName", "RemoveDistributionGroup");
                        ex.Data.Add("Parameters", parameters);

                        throw new Exception(error.ToString(), ex);
                    }
                }
            }
            catch (SystemException ex)
            {
                List<KeyValuePair<string, string>> parameters = new List<KeyValuePair<string, string>>()
                {
                    new KeyValuePair<string, string>("Identity", groupName),
                    new KeyValuePair<string, string>("Confirm", false.ToString())
                };

                ExchangeCommandException exc = new ExchangeCommandException();
                ex.Data.Add("CommandName", "remove-distributiongroup");
                ex.Data.Add("MethodName", "RemoveDistributionGroup");
                ex.Data.Add("Parameters", parameters);

                throw new Exception(ex.Message, exc);

            }
            finally
            {
                pl.Error.Close();
                pl.Stop(); pl.Dispose();
            }

        }

        public virtual bool RemoveDistributionListMember(string dlname, string memberidentity)
        {
            bool rc = true;

            Command cmd = new Command("remove-distributiongroupmember");
            cmd.Parameters.Add("Identity", dlname);
            cmd.Parameters.Add("Member", memberidentity);
            cmd.Parameters.Add("Confirm", false);

            Pipeline pl = rs.CreatePipeline();
            pl.Commands.Add(cmd);

            try
            {
                Collection<PSObject> results = pl.Invoke();
                if (pl.Error.Count > 0)
                {
                    var error = pl.Error.Read();
                    if (error != null)
                    {
                        List<KeyValuePair<string, string>> parameters = new List<KeyValuePair<string, string>>()
                        {
                            new KeyValuePair<string, string>("Identity", dlname),
                            new KeyValuePair<string, string>("Member", memberidentity),
                            new KeyValuePair<string, string>("Confirm", false.ToString())
                        };

                        ExchangeCommandException ex = new ExchangeCommandException();
                        ex.Data.Add("CommandName", "remove-distributiongroupmember");
                        ex.Data.Add("MethodName", "RemoveDistributionListMember");
                        ex.Data.Add("Parameters", parameters);

                        throw new Exception(error.ToString(), ex);
                    }
                    rc = false;
                }
            }
            catch (SystemException ex)
            {
                List<KeyValuePair<string, string>> parameters = new List<KeyValuePair<string, string>>()
                {
                    new KeyValuePair<string, string>("Identity", dlname),
                    new KeyValuePair<string, string>("Member", memberidentity),
                    new KeyValuePair<string, string>("Confirm", false.ToString())
                };

                ExchangeCommandException exc = new ExchangeCommandException();
                ex.Data.Add("CommandName", "remove-distributiongroupmember");
                ex.Data.Add("MethodName", "RemoveDistributionListMember");
                ex.Data.Add("Parameters", parameters);

                throw new Exception(ex.Message, exc);

            }
            finally
            {
                pl.Error.Close();
                pl.Stop(); pl.Dispose();
            }

            return rc;

        }

        public virtual void RemoveMailContact(string identity)
        {
            Command cmd = new Command("remove-mailcontact");
            cmd.Parameters.Add("Identity", identity);
            cmd.Parameters.Add("Confirm", false);

            Pipeline pl = rs.CreatePipeline();
            pl.Commands.Add(cmd);

            try
            {
                Collection<PSObject> results = pl.Invoke();
                if (pl.Error.Count > 0)
                {
                    var error = pl.Error.Read();
                    if (error != null)
                    {
                        List<KeyValuePair<string, string>> parameters = new List<KeyValuePair<string, string>>()
                        {
                            new KeyValuePair<string, string>("Identity", identity),
                            new KeyValuePair<string, string>("Confirm", false.ToString())
                        };

                        ExchangeCommandException ex = new ExchangeCommandException();
                        ex.Data.Add("CommandName", "remove-mailcontact");
                        ex.Data.Add("MethodName", "RemoveMailContact");
                        ex.Data.Add("Parameters", parameters);

                        throw new Exception(error.ToString(), ex);
                    }
                }
            }
            catch (SystemException ex)
            {
                List<KeyValuePair<string, string>> parameters = new List<KeyValuePair<string, string>>()
                {
                    new KeyValuePair<string, string>("Identity", identity),
                    new KeyValuePair<string, string>("Confirm", false.ToString())
                };

                ExchangeCommandException exc = new ExchangeCommandException();
                ex.Data.Add("CommandName", "remove-mailcontact");
                ex.Data.Add("MethodName", "RemoveMailContact");
                ex.Data.Add("Parameters", parameters);

                throw new Exception(ex.Message, exc);
            }
            finally
            {
                pl.Error.Close();
                pl.Stop(); pl.Dispose();
            }

        }

        public virtual bool UpdateMailContact(string identity, string displayname)
        {
            bool rc = true;

            if (!string.IsNullOrEmpty(displayname))
            {
                displayname = ExchangeHelper.TrimLength(displayname, 64);
                Command cmd = new Command("set-mailcontact");
                cmd.Parameters.Add("Identity", identity);
                cmd.Parameters.Add("DisplayName", displayname);
                cmd.Parameters.Add("Confirm", false);
                cmd.Parameters.Add("ForceUpgrade");

                Pipeline pl = rs.CreatePipeline();
                pl.Commands.Add(cmd);

                try
                {
                    Collection<PSObject> results = pl.Invoke();
                    if (pl.Error.Count > 0)
                    {
                        var error = pl.Error.Read();
                        if (error != null)
                        {
                            List<KeyValuePair<string, string>> parameters = new List<KeyValuePair<string, string>>()
                            {
                                new KeyValuePair<string, string>("Identity", identity),
                                new KeyValuePair<string, string>("DisplayName", displayname),
                                new KeyValuePair<string, string>("ForceUpgrade", "")
                            };

                            ExchangeCommandException ex = new ExchangeCommandException();
                            ex.Data.Add("CommandName", "set-mailcontact");
                            ex.Data.Add("MethodName", "UpdateMailContact");
                            ex.Data.Add("Parameters", parameters);

                            throw new Exception(error.ToString(), ex);
                        }


                        rc = false;
                    }
                }
                catch (SystemException ex)
                {
                    rc = false;
                    List<KeyValuePair<string, string>> parameters = new List<KeyValuePair<string, string>>()
                    {
                        new KeyValuePair<string, string>("Identity", identity),
                        new KeyValuePair<string, string>("DisplayName", displayname),
                        new KeyValuePair<string, string>("ForceUpgrade", "")
                    };

                    ExchangeCommandException exc = new ExchangeCommandException();
                    ex.Data.Add("CommandName", "set-mailcontact");
                    ex.Data.Add("MethodName", "UpdateMailContact");
                    ex.Data.Add("Parameters", parameters);

                    throw new Exception(ex.Message, exc);


                }
                finally
                {
                    pl.Error.Close();
                    pl.Stop(); pl.Dispose();
                }
            }
            else
            {
                rc = false;
                Trace.AddLog(EventLevel.Warning, DateTime.Now, "UpdateMailContact", "Exchange", "Update did not occur. Display Name had no value", String.Format("Parameters Identity: {1}, DisplayName: {2}", "set-mailcontact", identity, displayname));
            }
            return rc;

        }

        public virtual bool UpdateMailContact(string identity, string[] sparams, string[] values)
        {
            bool rc = true;
            Command cmd = new Command("set-contact");
            cmd.Parameters.Add("Identity", identity);

            for (int i = 0; i < sparams.Length; i++)
            {
                if (!string.IsNullOrEmpty(values[i]))
                    cmd.Parameters.Add(sparams[i], values[i]);
            }

           Pipeline pl = rs.CreatePipeline();
            pl.Commands.Add(cmd);

            try
            {
                Collection<PSObject> results = pl.Invoke();
                if (pl.Error.Count > 0)
                {
                    var error = pl.Error.Read();
                    if (error != null)
                    {
                        List<KeyValuePair<string, string>> parameters = new List<KeyValuePair<string, string>>();
                        parameters.Add(new KeyValuePair<string, string>("Identity", identity));
                        for (int i = 0; i < sparams.Length; i++)
                        {
                            if (!string.IsNullOrEmpty(values[i]))
                                cmd.Parameters.Add(sparams[i], values[i]);
                        }

                        ExchangeCommandException ex = new ExchangeCommandException();
                        ex.Data.Add("CommandName", "set-contact");
                        ex.Data.Add("MethodName", "UpdateMailContact");
                        ex.Data.Add("Parameters", parameters);

                        throw new Exception(error.ToString(), ex);
                    }
                    rc = false;
                }
            }
            catch (SystemException ex)
            {
                rc = false;
                List<KeyValuePair<string, string>> parameters = new List<KeyValuePair<string, string>>();
                parameters.Add(new KeyValuePair<string, string>("Identity", identity));
                for (int i = 0; i < sparams.Length; i++)
                {
                    if (!string.IsNullOrEmpty(values[i]))
                        cmd.Parameters.Add(sparams[i], values[i]);
                }

                ExchangeCommandException exc = new ExchangeCommandException();
                ex.Data.Add("CommandName", "set-contact");
                ex.Data.Add("MethodName", "UpdateMailContact");
                ex.Data.Add("Parameters", parameters);

                throw new Exception(ex.Message, exc);

            }
            finally
            {
                pl.Error.Close();
                pl.Stop(); pl.Dispose();
            }

            return rc;

        }

        public virtual bool UpdateMailContact(string identity, string displayname, string emailaddress)
        {
            bool rc = true;
            Command cmd = new Command("set-mailcontact");
            cmd.Parameters.Add("Identity", identity);
            if (!string.IsNullOrEmpty(displayname))
            {
                cmd.Parameters.Add("DisplayName", displayname);
            }

            // Only Update the External Email Address when the identity is not an email address
            if (!identity.IsValidEmail())
            {
                cmd.Parameters.Add("ExternalEmailAddress", emailaddress);
                cmd.Parameters.Add("Confirm", false);
                cmd.Parameters.Add("ForceUpgrade");
            }

            Pipeline pl = rs.CreatePipeline();
            pl.Commands.Add(cmd);

            try
            {
                Collection<PSObject> results = pl.Invoke();
                if (pl.Error.Count > 0)
                {
                    var error = pl.Error.Read();
                    if (error != null)
                    {
                        List<KeyValuePair<string, string>> parameters = new List<KeyValuePair<string, string>>()
                        {
                            new KeyValuePair<string, string>("Identity", identity),
                            new KeyValuePair<string, string>("DisplayName", displayname),
                            new KeyValuePair<string, string>("ExternalEmailAddress", emailaddress),
                            new KeyValuePair<string, string>("ForceUpgrade", "")
                        };

                        ExchangeCommandException ex = new ExchangeCommandException();
                        ex.Data.Add("CommandName", "set-mailcontact");
                        ex.Data.Add("MethodName", "UpdateMailContact");
                        ex.Data.Add("Parameters", parameters);

                        throw new Exception(error.ToString(), ex);
                    }
                    rc = false;
                }
            }
            catch (SystemException ex)
            {
                rc = false;
                List<KeyValuePair<string, string>> parameters = new List<KeyValuePair<string, string>>()
                {
                    new KeyValuePair<string, string>("Identity", identity),
                    new KeyValuePair<string, string>("DisplayName", displayname),
                    new KeyValuePair<string, string>("ExternalEmailAddress", emailaddress),
                    new KeyValuePair<string, string>("ForceUpgrade", "")
                };

                ExchangeCommandException exc = new ExchangeCommandException();
                ex.Data.Add("CommandName", "set-mailcontact");
                ex.Data.Add("MethodName", "UpdateMailContact");
                ex.Data.Add("Parameters", parameters);

                throw new Exception(ex.Message, exc);
            }
            finally
            {
                pl.Error.Close();
                pl.Stop(); pl.Dispose();
            }
            return rc;

        }

        public virtual bool UpdateMailContact(string identity, string alias, string name, string displayname, string emailaddress)
        {
            bool rc = true;
            Command cmd = new Command("set-mailcontact");
            cmd.Parameters.Add("Identity", identity);
            cmd.Parameters.Add("Alias", alias);
            cmd.Parameters.Add("Name", name);
            if (!string.IsNullOrEmpty(displayname))
            {
                cmd.Parameters.Add("DisplayName", displayname);
            }
            cmd.Parameters.Add("ExternalEmailAddress", emailaddress);
            cmd.Parameters.Add("Confirm", false);
            cmd.Parameters.Add("ForceUpgrade");

            Pipeline pl = rs.CreatePipeline();
            pl.Commands.Add(cmd);

            try
            {
                Collection<PSObject> results = pl.Invoke();
                if (pl.Error.Count > 0)
                {
                    var error = pl.Error.Read();
                    if (error != null)
                    {
                        List<KeyValuePair<string, string>> parameters = new List<KeyValuePair<string, string>>()
                        {
                            new KeyValuePair<string, string>("Identity", identity),
                            new KeyValuePair<string, string>("Alias", alias),
                            new KeyValuePair<string, string>("Name", name),
                            new KeyValuePair<string, string>("DisplayName", displayname),
                            new KeyValuePair<string, string>("ExternalEmailAddress", emailaddress),
                            new KeyValuePair<string, string>("ForceUpgrade", "")
                        };

                        ExchangeCommandException ex = new ExchangeCommandException();
                        ex.Data.Add("CommandName", "set-mailcontact");
                        ex.Data.Add("MethodName", "UpdateMailContact");
                        ex.Data.Add("Parameters", parameters);

                        throw new Exception(error.ToString(), ex);
                    }
                    rc = false;
                }
            }
            catch (SystemException ex)
            {
                rc = false;
                List<KeyValuePair<string, string>> parameters = new List<KeyValuePair<string, string>>()
                {
                    new KeyValuePair<string, string>("Identity", identity),
                    new KeyValuePair<string, string>("Alias", alias),
                    new KeyValuePair<string, string>("Name", name),
                    new KeyValuePair<string, string>("DisplayName", displayname),
                    new KeyValuePair<string, string>("ExternalEmailAddress", emailaddress),
                    new KeyValuePair<string, string>("ForceUpgrade", "")
                };

                ExchangeCommandException exc = new ExchangeCommandException();
                ex.Data.Add("CommandName", "set-mailcontact");
                ex.Data.Add("MethodName", "UpdateMailContact");
                ex.Data.Add("Parameters", parameters);

                throw new Exception(ex.Message, exc);

            }
            finally
            {
                pl.Error.Close();
                pl.Stop(); pl.Dispose();
            }
            return rc;

        }

        /// <summary>
        /// Not for use with Exchange Online
        /// </summary>
        /// <param name="identity"></param>
        /// <param name="alias"></param>
        /// <returns></returns>
        public virtual bool UpdateMailContactAlias(string identity, string alias)
        {
            bool rc = true;
            Command cmd = new Command("set-mailcontact");
            cmd.Parameters.Add("Identity", identity);
            // cmd.Parameters.Add("Alias", alias);

            if (alias.IsValidEmail())
                cmd.Parameters.Add("Alias", ExchangeHelper.GenerateNameFromEmailAddress(alias));
            else
                cmd.Parameters.Add("Alias", alias);
            cmd.Parameters.Add("Confirm", false);
            cmd.Parameters.Add("ForceUpgrade");

            Pipeline pl = rs.CreatePipeline();
            pl.Commands.Add(cmd);

            try
            {
                Collection<PSObject> results = pl.Invoke();
                if (pl.Error.Count > 0)
                {
                    var error = pl.Error.Read();
                    if (error != null)
                    {
                        List<KeyValuePair<string, string>> parameters = new List<KeyValuePair<string, string>>()
                        {
                            new KeyValuePair<string, string>("Identity", identity),
                            new KeyValuePair<string, string>("Alias", alias),
                            new KeyValuePair<string, string>("ForceUpgrade", string.Empty)
                        };

                        ExchangeCommandException ex = new ExchangeCommandException();
                        ex.Data.Add("CommandName", "set-mailcontact");
                        ex.Data.Add("MethodName", "UpdateMailContactAlias");
                        ex.Data.Add("Parameters", parameters);

                        throw new Exception(error.ToString(), ex);
                    }

                    rc = false;
                }
            }
            catch (SystemException ex)
            {
                rc = false;
                List<KeyValuePair<string, string>> parameters = new List<KeyValuePair<string, string>>()
                {
                    new KeyValuePair<string, string>("Identity", identity),
                    new KeyValuePair<string, string>("Alias", alias),
                    new KeyValuePair<string, string>("ForceUpgrade", string.Empty)
                };

                ExchangeCommandException exc = new ExchangeCommandException();
                ex.Data.Add("CommandName", "set-mailcontact");
                ex.Data.Add("MethodName", "UpdateMailContactAlias");
                ex.Data.Add("Parameters", parameters);

                throw new Exception(ex.Message, exc);

            }
            finally
            {
                pl.Error.Close();
                pl.Stop(); pl.Dispose();
            }
            return rc;

        }

        public virtual bool UpdateMailUser(string identity, string[] sparams, string[] values)
        {
            bool rc = true;
            Command cmd = new Command("set-mailbox");
            cmd.Parameters.Add("Identity", identity);

            for (int i = 0; i < sparams.Length; i++)
            {
                cmd.Parameters.Add(sparams[i], values[i]);
            }

            Pipeline pl = rs.CreatePipeline();
            pl.Commands.Add(cmd);

            try
            {
                Collection<PSObject> results = pl.Invoke();
                if (pl.Error.Count > 0)
                {

                    var error = pl.Error.Read();
                    if (error != null)
                    {
                        List<KeyValuePair<string, string>> parameters = new List<KeyValuePair<string, string>>();
                        parameters.Add(new KeyValuePair<string, string>("Identity", identity));
                        for (int i = 0; i < sparams.Length; i++)
                        {
                            if (!string.IsNullOrEmpty(values[i]))
                                cmd.Parameters.Add(sparams[i], values[i]);
                        }

                        ExchangeCommandException ex = new ExchangeCommandException();
                        ex.Data.Add("CommandName", "set-mailbox");
                        ex.Data.Add("MethodName", "UpdateMailUser");
                        ex.Data.Add("Parameters", parameters);

                        throw new Exception(error.ToString(), ex);
                    }
                    rc = false;
                }
            }
            catch (SystemException ex)
            {
                rc = false;
                List<KeyValuePair<string, string>> parameters = new List<KeyValuePair<string, string>>();
                parameters.Add(new KeyValuePair<string, string>("Identity", identity));
                for (int i = 0; i < sparams.Length; i++)
                {
                    if (!string.IsNullOrEmpty(values[i]))
                        cmd.Parameters.Add(sparams[i], values[i]);
                }

                ExchangeCommandException exc = new ExchangeCommandException();
                ex.Data.Add("CommandName", "set-mailbox");
                ex.Data.Add("MethodName", "UpdateMailUser");
                ex.Data.Add("Parameters", parameters);

                throw new Exception(ex.Message, exc);
            }
            finally
            {
                pl.Error.Close();
                pl.Stop(); pl.Dispose();
            }
            return rc;

        }

        public virtual bool UpdateMailUser(string identity, string displayname, string emailaddress)
        {
            bool rc = true;
            Command cmd = new Command("set-mailbox");
            cmd.Parameters.Add("Identity", identity);
            if (!string.IsNullOrEmpty(displayname))
            {
                cmd.Parameters.Add("DisplayName", displayname);
            }
            cmd.Parameters.Add("ExternalEmailAddress", emailaddress);
            cmd.Parameters.Add("Force");

            Pipeline pl = rs.CreatePipeline();
            pl.Commands.Add(cmd);

            try
            {
                Collection<PSObject> results = pl.Invoke();
                if (pl.Error.Count > 0)
                {
                    var error = pl.Error.Read();
                    if (error != null)
                    {
                        List<KeyValuePair<string, string>> parameters = new List<KeyValuePair<string, string>>()
                        {
                            new KeyValuePair<string, string>("Identity", identity),
                            new KeyValuePair<string, string>("DisplayName", displayname),
                            new KeyValuePair<string, string>("ExternalEmailAddress", emailaddress),
                            new KeyValuePair<string, string>("Force", string.Empty)
                        };

                        ExchangeCommandException ex = new ExchangeCommandException();
                        ex.Data.Add("CommandName", "set-mailbox");
                        ex.Data.Add("MethodName", "UpdateMailUser");
                        ex.Data.Add("Parameters", parameters);

                        throw new Exception(error.ToString(), ex);
                    }
                    rc = false;
                }
            }
            catch (SystemException ex)
            {
                rc = false;
                List<KeyValuePair<string, string>> parameters = new List<KeyValuePair<string, string>>()
                {
                    new KeyValuePair<string, string>("Identity", identity),
                    new KeyValuePair<string, string>("DisplayName", displayname),
                    new KeyValuePair<string, string>("ExternalEmailAddress", emailaddress),
                    new KeyValuePair<string, string>("Force", string.Empty)
                };

                ExchangeCommandException exc = new ExchangeCommandException();
                ex.Data.Add("CommandName", "set-mailbox");
                ex.Data.Add("MethodName", "UpdateMailUser");
                ex.Data.Add("Parameters", parameters);

                throw new Exception(ex.Message, exc);
            }
            finally
            {
                pl.Error.Close();
                pl.Stop(); pl.Dispose();
            }
            return rc;

        }

        public virtual void CloseRunspaces()
        {

        }

        public virtual bool UserExists(string emailaddress)
        {
            bool rc = false;

            Command cmd = new Command("get-mailbox");
            cmd.Parameters.Add("Identity", emailaddress);

            Pipeline pl = rs.CreatePipeline();
            pl.Commands.Add(cmd);

            try
            {
                Collection<PSObject> results = pl.Invoke();
                if (pl.Error.Count > 0)
                {
                    var error = pl.Error.Read();
                    if (error != null)
                    {
                        List<KeyValuePair<string, string>> parameters = new List<KeyValuePair<string, string>>()
                        {
                            new KeyValuePair<string, string>("Identity", emailaddress)
                        };

                        ExchangeCommandException ex = new ExchangeCommandException();
                        ex.Data.Add("CommandName", "get-mailbox");
                        ex.Data.Add("MethodName", "UserExists");
                        ex.Data.Add("Parameters", parameters);

                        throw new Exception(error.ToString(), ex);
                    }
                    rc = false;
                }

                if (results.Count > 0)
                {
                    rc = true;
                    if (results.Count > 1)
                    {
                        Trace.AddLog(EventLevel.Warning, DateTime.Now, "UserExists", "Exchange", "Multiple Results for User with Email Address " + emailaddress, string.Empty);
                    }
                }

            }
            catch (SystemException ex)
            {
                List<KeyValuePair<string, string>> parameters = new List<KeyValuePair<string, string>>()
                {
                    new KeyValuePair<string, string>("Identity", emailaddress)
                };

                ExchangeCommandException exc = new ExchangeCommandException();
                ex.Data.Add("CommandName", "get-mailbox");
                ex.Data.Add("MethodName", "UserExists");
                ex.Data.Add("Parameters", parameters);

                throw new Exception(ex.Message, exc);
            }
            finally
            {
                pl.Error.Close();
                pl.Stop(); pl.Dispose();
            }

            return rc;
        }
    }
}
