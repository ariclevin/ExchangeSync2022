using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Net;
using System.Net.NetworkInformation;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace ExchangeSync
{
    public static class Extensions
    {
        public static string ToPascalCase(this string s)
        {
            var cultureInfo = System.Threading.Thread.CurrentThread.CurrentCulture;
            return cultureInfo.TextInfo.ToTitleCase(s.ToLower());

        }
        public static string ToPascalCase(this string s, string cultureInfoName)
        {
            var cultureInfo = new CultureInfo(cultureInfoName);
            return cultureInfo.TextInfo.ToTitleCase(s.ToLower());
        }

        public static int ToInt(this Enum e)
        {
            return Convert.ToInt32(e);
        }

        public static string StripPunctuation(this string s)
        {
            var sb = new StringBuilder();
            foreach (char c in s)
            {
                if (char.IsLetterOrDigit(c))
                    sb.Append(c);

            }
            return sb.ToString();
        }

        public static Single ParseVersion(this string input)
        {
            if (input.EndsWith("."))
                input = input.Substring(0, input.Length - 1);

            try
            {
                Version ver = Version.Parse(input);
                string versionNumber = ver.Major.ToString() + "." + ver.Minor.ToString();
                return Convert.ToSingle(versionNumber);
            }
            
            catch (ArgumentNullException)
            {
                Console.WriteLine("Error: String to be parsed is null.");
                return Convert.ToSingle("0.0");
            }
            catch (ArgumentOutOfRangeException)
            {
                Console.WriteLine("Error: Negative value in '{0}'.", input);
                return Convert.ToSingle("0.0");
            }
            catch (ArgumentException)
            {
                Console.WriteLine("Error: Bad number of components in '{0}'.", input);
                return Convert.ToSingle("0.0");
            }
            catch (FormatException)
            {
                Console.WriteLine("Error: Non-integer value in '{0}'.", input);
                return Convert.ToSingle("0.0");
            }
            catch (OverflowException)
            {
                Console.WriteLine("Error: Number out of range in '{0}'.", input);
                return Convert.ToSingle("0.0");
            }
        }

        public static string MaxLength(this string s, int maxLength)
        {
            string rc = "";
            if (!string.IsNullOrEmpty(s))
            {
                if (s.Length > maxLength)
                    rc = s.Substring(0, maxLength);
                else
                    rc = s;
            }
            return rc;
        }

        private static bool IsValidEmailOld(this string emailAddress)
        {
            const string MatchEmailPattern =
                    @"^(([\w-]+\.)+[\w-]+|([a-zA-Z]{1}|[\w-]{2,}))@"
             + @"((([0-1]?[0-9]{1,2}|25[0-5]|2[0-4][0-9])\.([0-1]?[0-9]{1,2}|25[0-5]|2[0-4][0-9])\."
             + @"([0-1]?[0-9]{1,2}|25[0-5]|2[0-4][0-9])\.([0-1]?[0-9]{1,2}|25[0-5]|2[0-4][0-9])){1}|"
             + @"([a-zA-Z]+[\w-]+\.)+[a-zA-Z]{2,4})$";

            bool rc = false;
            if (!string.IsNullOrEmpty(emailAddress))
            {
                rc = Regex.IsMatch(emailAddress, MatchEmailPattern);
            }

            return rc;
        }

        public static bool IsValidEmail(this string emailAddress)
        {
            bool rc = false;

            const string MatchEmailPattern = @"^(?("")("".+?(?<!\\)""@)|(([0-9a-z]((\.(?!\.))|[-!#\$%&'\*\+/=\?\^`\{\}\|~\w])*)(?<=[0-9a-z])@))" +
                      @"(?(\[)(\[(\d{1,3}\.){3}\d{1,3}\])|(([0-9a-z][-\w]*[0-9a-z]*\.)+[a-z0-9][\-a-z0-9]{0,22}[a-z0-9]))$";

            if (String.IsNullOrEmpty(emailAddress))
                return false;

            // Return true if strIn is in valid e-mail format.
            try
            {
                rc = Regex.IsMatch(emailAddress, MatchEmailPattern, RegexOptions.IgnoreCase);
            }
            catch (Exception ex)
            {
                rc = false;
            }

            return rc;
        }

        public static bool IsEmailInDomain(this string emailAddress, string domainName)
        {
            bool rc = false;

            if (domainName.Contains(';'))
            {
                string[] domains = domainName.Split(';');
                foreach (string domain in domains)
                {
                    if (emailAddress.Contains(domain))
                    {
                        rc = true;
                        break;
                    }
                }
            }
            else
            {
                rc = emailAddress.Contains(domainName);
            }

            return rc;
        }


        public static string ToPlainString(this Enum value)
        {
            string rc = value.ToString();
            rc = rc.Replace('_', ' ');
            return rc;
        }

        public static bool IsValidUrl(this string url)
        {
            const string MatchUrlPattern = "(([a-zA-Z][0-9a-zA-Z+\\-\\.]*:)?/{0,2}[0-9a-zA-Z;/?:@&=+$\\.\\-_!~*'()%]+)?(#[0-9a-zA-Z;/?:@&=+$\\.\\-_!~*'()%]+)?";

            bool rc = false;

            if (!string.IsNullOrEmpty(url))
            {
                rc = Regex.IsMatch(url, MatchUrlPattern);
            }

            return rc;
        }

        public static bool Contains(this List<KeyValuePair<Guid, string>> list, Guid key)
        {
            return list.Any(c => c.Key == key);
        }

        public static bool Contains(this List<AssemblyVersions> list, string assemblyName)
        {
            return list.Any(c => c.AssemblyName == assemblyName);
        }

        public static string GetValue(this KeyValuePair<Guid, string> kvp)
        {
            if (!string.IsNullOrEmpty(kvp.Value))
                return kvp.Value;
            else
                return string.Empty;

        }
    }

    
}
