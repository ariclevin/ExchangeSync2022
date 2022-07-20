using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace ExchangeSync
{
    public static class ExchangeHelper
    {
        public static string TrimSpaces(string value)
        {
            string rc = "";
            rc = value.Replace(".", "");
            rc = rc.Replace(",", "");
            rc = rc.Replace("'", "");
            rc = rc.Replace("\"", "");
            rc = rc.Replace("(", "");
            rc = rc.Replace(")", "");
            rc = rc.Replace(" ", "");

            return rc;
        }

        public static string ValidateAlias(string value)
        {
            string rc = string.Empty;

            if (value.Length > 0)
            {
                foreach (char c in value)
                {
                    if (char.IsLetterOrDigit(c))
                        rc += c.ToString();
                    else if (char.IsWhiteSpace(c))
                        rc += '_';
                    else
                    {
                        switch (c)
                        {
                            case '!':
                            case '#':
                            case '$':
                            case '%':
                            case '\'':
                            case '*':
                            case '+':
                            case '-':
                            case '/':
                            case '=':
                            case '?':
                            case '^':
                            case '_':
                            case '{':
                            case '}':
                            case '|':
                            case '~':
                            case '`':
                            case '.':
                                rc += c.ToString();
                                break;
                            default:
                                break;
                        }
                    }
                }
            }
            return rc;
        }

        /// <summary>
        /// Trims the length of a string
        /// </summary>
        public static string TrimLength(string value, int length)
        {
            if (!string.IsNullOrEmpty(value))
            {
                if (value.Length > 0)
                {
                    if (value.Length > length)
                        value = value.Substring(0, length);
                }
            }
            else
                value = string.Empty;

            return value;
        }

        public static string TrimInvalidCharacters(string value)
        {
            string rc = "";
            rc = value.TrimStart(' ');
            rc = value.TrimEnd(' ');

            rc = Encoding.Unicode.GetString(Encoding.Unicode.GetBytes(rc));
            return rc;
        }

        public static string FormatName(string firstName, string lastName, string custom = "")
        {
            string rc = "";

            string formatType = AppSetting.ExchangeNameFieldFormat.Value;
            string formatValue = AppSetting.ExchangeNameFieldValue.Value;

            if (!string.IsNullOrEmpty(formatValue))
            {
                if (formatValue.Contains("#"))
                {
                    formatValue = formatValue.Replace("#", custom);
                }
            }

            if (!string.IsNullOrEmpty(formatType))
            {
                switch (formatType.ToUpper())
                {
                    case "PREFIX":
                        rc = formatValue + firstName + " " + lastName;
                        break;
                    case "SUFFIX":
                        rc = firstName + " " + lastName + formatValue;
                        break;
                    case "SEPARATOR":
                        rc = firstName + formatValue + lastName;
                        break;
                    default:
                        rc = firstName + " " + lastName;
                        break;
                }
            }
            else
            {
                rc = firstName + " " + lastName;
            }

            if (rc.StartsWith(" "))
                rc = rc.TrimStart(' ');

            rc = TrimInvalidCharacters(rc);
            if (rc.Length > 64)
                rc = rc.Substring(0, 64);
            return rc;
        }

        public static string FormatName(string firstName, string lastName, string emailAddress, string custom = "")
        {
            string rc = "";
            string formatType = AppSetting.ExchangeNameFieldFormat.Value;
            string formatValue = AppSetting.ExchangeNameFieldValue.Value;

            if (!string.IsNullOrEmpty(formatValue))
            {
                if (formatValue.Contains("#"))
                {
                    formatValue = formatValue.Replace("#", custom);
                }
            }

            if (!string.IsNullOrEmpty(formatType))
            {
                switch (formatType.ToUpper())
                {
                    case "PREFIX":
                        rc = formatValue + firstName + " " + lastName;
                        break;
                    case "SUFFIX":
                        rc = firstName + " " + lastName + formatValue;
                        break;
                    case "SEPARATOR":
                        rc = firstName + formatValue + lastName;
                        break;
                    case "APPENDDOMAIN":
                        rc = firstName + lastName + "_" + GetEmailDomainPrefixFromEmailAddress(emailAddress);
                        break;
                    default:
                        rc = firstName + " " + lastName;
                        break;
                }
            }
            else
            {
                rc = firstName + " " + lastName;
            }

            if (rc.StartsWith(" "))
                rc = rc.TrimStart(' ');

            rc = TrimInvalidCharacters(rc);
            if (rc.Length > 64)
                rc = rc.Substring(0, 64);
            return rc;

        }

        public static string FormatAlias(string firstName, string lastName, string custom = "")
        {
            string rc = "";
            string formatType = AppSetting.ExchangeAliasFieldFormat.Value;
            string formatValue = AppSetting.ExchangeAliasFieldValue.Value;

            firstName = TrimSpaces(firstName); lastName = TrimSpaces(lastName);

            if (!string.IsNullOrEmpty(formatValue))
            {
                if (formatValue.Contains("#"))
                {
                    formatValue = formatValue.Replace("#", custom);
                }
            }

            if (!string.IsNullOrEmpty(formatType))
            {
                switch (formatType.ToUpper())
                {
                    case "PREFIX":
                        rc = formatValue + firstName + lastName;
                        break;
                    case "SUFFIX":
                        rc = firstName + lastName + formatValue;
                        break;
                    case "SEPARATOR":
                        rc = firstName + formatValue + lastName;
                        break;
                    default:
                        rc = firstName + lastName;
                        break;
                }
            }
            else
            {
                rc = firstName + lastName;
            }

            if (rc.StartsWith(" "))
                rc = rc.TrimStart(' ');

            rc = TrimInvalidCharacters(rc);
            if (rc.Length > 64)
                rc = rc.Substring(0, 64);
            return rc;
        }

        public static string FormatAlias(string firstName, string lastName, string emailAddress, string custom = "")
        {
            string rc = "";
            string formatType = AppSetting.ExchangeAliasFieldFormat.Value;
            string formatValue = AppSetting.ExchangeAliasFieldValue.Value;

            firstName = TrimSpaces(firstName); lastName = TrimSpaces(lastName);

            if (!string.IsNullOrEmpty(formatValue))
            {
                if (formatValue.Contains("#"))
                {
                    formatValue = formatValue.Replace("#", custom);
                }
            }

            if (!string.IsNullOrEmpty(formatType))
            {
                switch (formatType.ToUpper())
                {
                    case "PREFIX":
                        rc = formatValue + firstName + lastName;
                        break;
                    case "SUFFIX":
                        rc = firstName + lastName + formatValue;
                        break;
                    case "SEPARATOR":
                        rc = firstName + formatValue + lastName;
                        break;
                    case "APPENDDOMAIN":
                        rc = firstName + lastName + "_" + GetEmailDomainPrefixFromEmailAddress(emailAddress);
                        break;
                    default:
                        rc = firstName + lastName;
                        break;
                }
            }
            else
            {
                rc = firstName + lastName;
            }

            if (rc.StartsWith(" "))
                rc = rc.TrimStart(' ');

            rc = TrimInvalidCharacters(rc);
            if (rc.Length > 64)
                rc = rc.Substring(0, 64);
            return rc;
        }

        public static string GetEmailDomainPrefixFromEmailAddress(string emailAddress)
        {
            string rc = string.Empty;
            rc = emailAddress.Substring(emailAddress.IndexOf('@') + 1);
            rc = rc.Substring(0, rc.IndexOf('.'));
            return rc.ToPascalCase();
        }

        public static string GenerateNameFromEmailAddress(string emailAddress)
        {
            string rc = "";
            bool capitalizeNext = true; // first character is capitalized
            for (int i = 0; i < emailAddress.Length; i++)
            {
                char current = emailAddress[i];
                switch (current)
                {
                    case '.':
                        capitalizeNext = true;
                        break;
                    case '@':
                        rc += "_";
                        capitalizeNext = true;
                        break;
                    default:
                        if (capitalizeNext)
                            rc += current.ToString().ToUpper();
                        else
                            rc += current.ToString();
                        capitalizeNext = false;
                        break;
                }
            }
            return rc;
        }

        public static string GenerateDisplayName(string appSettingValue, string fullname, string firstname, string lastname)
        {
            string rc = "";

            if (ExchangeSync.ConfigSetting.ExchangeServerVersion != ExchangeServerType.Exchange_2010)
            {
                switch (appSettingValue)
                {
                    case "NOUPDATE":
                        rc = string.Empty;
                        break;
                    case "CRMSETTING":
                        rc = fullname;
                        break;
                    case "FIRSTLAST":
                        rc = String.Format("{0} {1}", firstname, lastname);
                        break;
                    case "LASTFIRST":
                        rc = String.Format("{0}, {1}", lastname, firstname);
                        break;
                    default:
                        rc = fullname;
                        break;
                }
            }
            else
            {
                rc = fullname;
            }

            return rc;
        }

    }


}
