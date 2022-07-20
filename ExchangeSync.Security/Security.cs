using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;

namespace ExchangeSync
{
    public class Security
    {
        string _cryptoBytes = "";
        public string CryptoBytes
        {
            set
            {
                _cryptoBytes = value;
            }
            private get
            {
                return _cryptoBytes;
            }
        }
        public Security()
        { }

        public Security(string bytesString)
        {
            _cryptoBytes = bytesString;
        }

        public string Decrypt(string valueToDecrypt)
        {
            byte[] bytes = System.Text.Encoding.ASCII.GetBytes(CryptoBytes);
            if (String.IsNullOrEmpty(valueToDecrypt))
            {
                throw new ArgumentNullException
                   ("The string which needs to be decrypted can not be null.");
            }
            DESCryptoServiceProvider cryptoProvider = new DESCryptoServiceProvider();
            MemoryStream memoryStream = new MemoryStream
                    (Convert.FromBase64String(valueToDecrypt));
            CryptoStream cryptoStream = new CryptoStream(memoryStream,
                cryptoProvider.CreateDecryptor(bytes, bytes), CryptoStreamMode.Read);
            StreamReader reader = new StreamReader(cryptoStream);
            return reader.ReadToEnd();
        }

        public string Encrypt(string valueToEncrypt)
        {
            byte[] bytes = System.Text.Encoding.ASCII.GetBytes(CryptoBytes);
            if (String.IsNullOrEmpty(valueToEncrypt))
            {
                throw new ArgumentNullException
                       ("The string which needs to be encrypted can not be null.");
            }
            DESCryptoServiceProvider cryptoProvider = new DESCryptoServiceProvider();
            MemoryStream memoryStream = new MemoryStream();
            CryptoStream cryptoStream = new CryptoStream(memoryStream,
                cryptoProvider.CreateEncryptor(bytes, bytes), CryptoStreamMode.Write);
            StreamWriter writer = new StreamWriter(cryptoStream);
            writer.Write(valueToEncrypt);
            writer.Flush();
            cryptoStream.FlushFinalBlock();
            writer.Flush();
            return Convert.ToBase64String(memoryStream.GetBuffer(), 0, (int)memoryStream.Length);
        }

        public bool ValidateKey(string key)
        {
            const string KEY_FORMAT = "%#^^?-^^%%^-^^?^^-#%#%#-^^%#?";
            List<Int32> list = new List<Int32>();

            if (KEY_FORMAT.Length == key.Length)
            {
                for (int i = 0; i < key.Length; i++)
                {
                    string k = key.Substring(i, 1);
                    string f = KEY_FORMAT.Substring(i, 1);

                    char current = Convert.ToChar(k);
                    switch (f)
                    {
                        case "^": // Upper or Lowecase Letter
                            if (!char.IsLetter(current))
                                return false;
                            break;
                        case "#":
                            if (!char.IsDigit(current))
                                return false;
                            break;
                        case "?":
                            if (!char.IsLetterOrDigit(current))
                                return false;
                            break;
                        case "%":
                            list.Add(Convert.ToInt32(k));
                            break;
                        default:
                            if (!(k.ToLower() == f.ToLower()))
                                return false;
                            break;
                    }
                }

                int sum = 0;
                foreach (int val in list)
                {
                    sum += val;
                }

                if ((sum % 7) != 0)
                    return false;

                return true;
            }
            else
            {
                return false;
            }
        }

        public static string StringToHex(string input)
        {
            string rc = "";
            char[] values = input.ToCharArray();
            foreach (char letter in values)
            {
                int value = Convert.ToInt32(letter);
                string hexValue = String.Format("{0:X}", value);
                rc += hexValue;
            }
            return rc;
        }

        public static string HexToString(string input)
        {
            StringBuilder sb = new StringBuilder();
            for (int i = 0; i < input.Length; i += 2)
            {
                string hs = input.Substring(i, 2);
                sb.Append(Convert.ToChar(Convert.ToUInt32(hs, 16)));
            }
            return sb.ToString();
        }
    }
}
