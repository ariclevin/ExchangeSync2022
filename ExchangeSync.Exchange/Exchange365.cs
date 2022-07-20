using System;
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
    public class Exchange365 : ExchangeBase
    {
        public event ConnectionHandler Connector;
        public delegate void ConnectionHandler(ExchangeBase exch, ExchangeConnectEventArgs e);

        /*
        public string AliasPrefix
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

        public string SHELL_URI
        {
            get
            {
                return string.Empty;
            }
        }
        */

        // PowerShell psh;
        Runspace rs;
        WSManConnectionInfo connectionInfo;

        public override bool AddDistributionListMember(string dlname, string memberidentity)
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

        public override void ClearPowerShellCommands()
        {
            throw new NotImplementedException();
        }

        public override ExchangeConnectEventArgs Connect()
        {
            ExchangeConnectEventArgs e = Connect(string.Empty, ConfigSetting.ExchangeCredentials);
            return e;

        }

        public override ExchangeConnectEventArgs Connect(string exchangeServerUrl, NetworkCredential credentials)
        {
            System.Security.SecureString encryptedPassword = EncryptPassword(credentials.Password);
            PSCredential credential = new PSCredential(credentials.UserName, encryptedPassword);

            connectionInfo = GetConnectionInfo(credential);

            rs = RunspaceFactory.CreateRunspace(connectionInfo);
            
            RunspaceStateInfo stateInfo;
            ExchangeConnectEventArgs e = new ExchangeConnectEventArgs();
            try
            {
                rs.Open();

                stateInfo = rs.RunspaceStateInfo;
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

            }
            catch (System.Exception ex)
            {
                stateInfo = rs.RunspaceStateInfo;
                e.ConnectionState = stateInfo.State;
                e.ConnectionReason = stateInfo.Reason.Message;
                e.Result = false;
            }

            return e;
        }

        private WSManConnectionInfo GetConnectionInfo(PSCredential credential) //, out string[] servers)
        {
            WSManConnectionInfo connectionInfo = new WSManConnectionInfo(new Uri("https://outlook.office365.com/PowerShell-LiveID?PSVersion=2.0"), "http://schemas.microsoft.com/powershell/Microsoft.Exchange", credential);
            connectionInfo.AuthenticationMechanism = AuthenticationMechanism.Basic;
            connectionInfo.SkipCACheck = true;
            connectionInfo.SkipCNCheck = true;
            connectionInfo.MaximumConnectionRedirectionCount = 4;
            connectionInfo.OperationTimeout = 5 * 60 * 1000; // 5 minutes.
            connectionInfo.OpenTimeout = 2 * 60 * 1000; // 2 minute.
            connectionInfo.IdleTimeout = 3 * 60 * 1000; // 3 minute.
            connectionInfo.CancelTimeout = 1 * 60 * 1000; // 1 minute.

            return connectionInfo;
        }

        private System.Security.SecureString EncryptPassword(string password)
        {
            System.Security.SecureString secureString = new System.Security.SecureString();
            foreach (char c in password)
                secureString.AppendChar(c);

            return secureString;
        }

        public override string GetServerName()
        {
            return "OnMicrosoft.com";
        }

        public override bool ContactExists(string emailAddressOrAlias)
        {
            bool rc = false;

            Command cmd = new Command("get-mailcontact");
            cmd.Parameters.Add("Identity", emailAddressOrAlias);

            Pipeline pl = rs.CreatePipeline();
            pl.Commands.Add(cmd);

            try
            {
                Collection<PSObject> results = pl.Invoke();
                if (results.Count > 0)
                {
                    rc = true;
                    if (results.Count > 1)
                    {
                        Trace.AddLog(EventLevel.Warning, DateTime.Now, "ContactExists", "Exchange", "Multiple Results for Email Address/Alias " + emailAddressOrAlias, string.Empty);
                    }
                }
                else
                {
                    if (pl.Error.Count > 0)
                    {
                        var error = pl.Error.Read();
                        if (error != null)
                        {

                            if (error.ToString().Contains("couldn't be found"))
                            {
                                Trace.AddLog(EventLevel.Warning, DateTime.Now, "ContactExists", "Exchange", "No Results for Email Address/Alias " + emailAddressOrAlias, string.Empty);
                                rc = false;
                            }
                            else
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
                        }
                        rc = false;
                        // Trace.AddLog(EventLevel.Warning, DateTime.Now, "CreateMailContact", "Exchange", "An error was returned from the Powershell Error Record", String.Format("Errors returned from calling {0} with parameters Name: {1}, External Email Address: {2}", "new-mailcontact", name, emailaddress));
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

        public override bool CreateDistributionGroup(string dlname, string ou)
        {
            return CreateDistributionGroup(dlname, ou, GroupType.Distribution);
        }

        public override bool CreateDistributionGroup(string dlname, string ou, GroupType groupType)
        {
            bool rc = true;

            Command cmd = new Command("new-DistributionGroup");
            cmd.Parameters.Add("Name", dlname);
            cmd.Parameters.Add("DisplayName", dlname);
            string alias = ExchangeHelper.ValidateAlias(dlname);
            cmd.Parameters.Add("Alias", alias);
            cmd.Parameters.Add("Type", groupType.ToString());

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
                            new KeyValuePair<string, string>("Alias", dlname),
                            new KeyValuePair<string, string>("Type", groupType.ToString())

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
                    new KeyValuePair<string, string>("Type", groupType.ToString())

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

        public override bool CreateMailContact(string emailaddress, string ou)
        {
            bool rc = true;
            string name = ExchangeHelper.GenerateNameFromEmailAddress(emailaddress);

            Command cmd = new Command("new-mailcontact");
            cmd.Parameters.Add("Name", name);
            cmd.Parameters.Add("DisplayName", emailaddress);
            cmd.Parameters.Add("ExternalEmailAddress", emailaddress);
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

        public override bool CreateMailContact(string displayname, string externalEmailAddress, string firstname, string lastname, string ou, out string alias, string primarySmtpAddress = "")
        {
            bool rc = true;
            if (!string.IsNullOrEmpty(displayname))
                displayname = ExchangeHelper.TrimLength(displayname, 64);
            else
                displayname = externalEmailAddress;

            string fullname = ExchangeHelper.TrimLength(displayname, 64);
            string name = ExchangeHelper.FormatName(firstname, lastname, externalEmailAddress, string.Empty);
            alias = ExchangeHelper.FormatAlias(firstname, lastname, externalEmailAddress, string.Empty);

            Command cmd = new Command("new-mailcontact");
            cmd.Parameters.Add("Name", name);
            // cmd.Parameters.Add("DisplayName", displayname);
            cmd.Parameters.Add("DisplayName", fullname);
            cmd.Parameters.Add("ExternalEmailAddress", externalEmailAddress);
            // cmd.Parameters.Add("PrimarySmtpAddress", primarySmtpAddress == "" ? externalEmailAddress : primarySmtpAddress);
            cmd.Parameters.Add("FirstName", firstname);
            cmd.Parameters.Add("LastName", lastname);
            cmd.Parameters.Add("Alias", alias);

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
                            new KeyValuePair<string, string>("FirstName", firstname),
                            new KeyValuePair<string, string>("LastName", lastname),
                            new KeyValuePair<string, string>("Alias", alias)
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
                    new KeyValuePair<string, string>("FirstName", firstname),
                    new KeyValuePair<string, string>("LastName", lastname),
                    new KeyValuePair<string, string>("Alias", alias)
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

        public override bool CreateMailContact(string displayName, string externalEmailAddress, string firstName, string lastName, string ou, string custom, out string alias, string primarySmtpAddress = "")
        {
            bool rc = true;
            if (!string.IsNullOrEmpty(displayName))
                displayName = ExchangeHelper.TrimLength(displayName, 64);
            else
                displayName = externalEmailAddress;

            string name = ExchangeHelper.FormatName(firstName, lastName, externalEmailAddress, custom);
            alias = ExchangeHelper.FormatAlias(firstName, lastName, externalEmailAddress, custom);
            firstName = ExchangeHelper.TrimLength(firstName, 64);
            lastName = ExchangeHelper.TrimLength(lastName, 64);


            Command cmd = new Command("new-mailcontact");
            cmd.Parameters.Add("Name", name);
            cmd.Parameters.Add("DisplayName", displayName);
            cmd.Parameters.Add("ExternalEmailAddress", externalEmailAddress);
            // cmd.Parameters.Add("PrimarySmtpAddress", primarySmtpAddress == "" ? externalEmailAddress : primarySmtpAddress);
            cmd.Parameters.Add("FirstName", firstName);
            cmd.Parameters.Add("LastName", lastName);
            cmd.Parameters.Add("Alias", alias);

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
                            new KeyValuePair<string, string>("DisplayName", displayName),
                            new KeyValuePair<string, string>("ExternalEmailAddress", externalEmailAddress),
                            new KeyValuePair<string, string>("FirstName", firstName),
                            new KeyValuePair<string, string>("LastName", lastName),
                            new KeyValuePair<string, string>("Alias", alias)
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
                    new KeyValuePair<string, string>("Alias", alias)
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

        public override bool Disconnect()
        {
            rs.Close();
            rs.Dispose();
            return true;
        }

        public override bool DistributionGroupExists(string dlname)
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

        public override bool ExchangeUserExists(string identity, bool mailboxEnabled)
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

        public override List<DistributionGroup> GetAllDistributionListGroups()
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

        public override List<string> GetAllDistributionLists()
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

        public override List<OrganizationalUnit> GetAllOrganizationalUnits()
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

        public override List<ExchangeContact> GetDistributionListMembers(string dlname)
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

        public override ExchangeContact GetRecipient(string identity)
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

                        throw new ExchangeCommandException(error.ToString(), ex);
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

        public override ExchangeContact GetMailContact(string identity)
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

        public override ExchangeContact GetMailUser(string identity)
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

        public override bool HideMailContact(string identity, bool hidden)
        {
            bool rc = true;

            Command cmd = new Command("set-mailcontact");
            cmd.Parameters.Add("Identity", identity);
            cmd.Parameters.Add("HiddenFromAddressListsEnabled", true);

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

        public override bool IsOUValid(string ouName)
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

        public override void RemoveAllDistributionListMembers(string dlname)
        {
            List<ExchangeContact> members = GetDistributionListMembers(dlname);
            Trace.AddLog(EventLevel.Information, DateTime.Now, "RemoveAllDistributionListMembers", "Exchange", "Calling Remove All Distribution List Members", members.Count.ToString() + " members retrieved");
            if (members.Count > 0)
            {
                foreach (ExchangeContact contact in members)
                {
                    string identity = "";
                    try
                    {
                        if (!string.IsNullOrEmpty(contact.Alias))
                        {
                            identity = contact.Alias;
                            RemoveDistributionListMember(dlname, contact.Alias);
                        }
                        else if (!string.IsNullOrEmpty(contact.PrimarySmtpAddress))
                        {
                            identity = contact.PrimarySmtpAddress;
                            RemoveDistributionListMember(dlname, contact.PrimarySmtpAddress);
                        }
                        else
                        {
                            identity = contact.Identity;
                            RemoveDistributionListMember(dlname, contact.Identity);
                        }
                        if (ExchangeSync.AppSetting.LoggingCriteriaVerbose.GetValue().ToLower() == "yes")
                        {
                            List<KeyValuePair<string, string>> parameters = new List<KeyValuePair<string, string>>()
                            {
                                new KeyValuePair<string, string>("Identity", dlname),
                                new KeyValuePair<string, string>("Member", identity)
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

        public void RemoveDistributionGroup(string groupName)
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

        public override bool RemoveDistributionListMember(string dlname, string memberidentity)
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

        public override void RemoveMailContact(string identity)
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

        public override bool UpdateMailContact(string identity, string displayname)
        {
            bool rc = true;

            if (!string.IsNullOrEmpty(displayname))
            {
                displayname = ExchangeHelper.TrimLength(displayname, 64);
                Command cmd = new Command("set-mailcontact");
                cmd.Parameters.Add("Identity", identity);
                cmd.Parameters.Add("DisplayName", displayname);
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

        public override bool UpdateMailContact(string identity, string[] sparams, string[] values)
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

        public override bool UpdateMailContact(string identity, string displayname, string emailaddress)
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

        public override bool UpdateMailContact(string identity, string alias, string name, string displayname, string emailaddress)
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
        public override bool UpdateMailContactAlias(string identity, string alias)
        {
            bool rc = true;
            Command cmd = new Command("set-mailcontact");
            cmd.Parameters.Add("Identity", identity);
            // cmd.Parameters.Add("Alias", alias);

            if (alias.IsValidEmail())
                cmd.Parameters.Add("Alias", ExchangeHelper.GenerateNameFromEmailAddress(alias));
            else
                cmd.Parameters.Add("Alias", alias);
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

        public override bool UpdateMailUser(string identity, string[] sparams, string[] values)
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

        public override bool UpdateMailUser(string identity, string displayname, string emailaddress)
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

        public override bool UserExists(string emailaddress)
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

        public override void CloseRunspaces()
        {

        }


    }
}
