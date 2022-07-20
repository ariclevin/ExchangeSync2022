using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using LogicNP.CryptoLicensing;

namespace ExchangeSync
{
    public class CoreLicense
    {
        enum FeatureList
        {
            None = 0,
            Windows = 1,
            AutoSync = 2 //,
            //ConsoleSync = 4,
        }

        public CryptoLicense license;

        public CoreLicense(string licenseCode)
        {
            license = new CryptoLicense();

            license.ValidationKey = "AMAAMADYLN6qRVjvUEWSzEw0UL1lN8bnWzRVTeB0YCgvKtr8mYotNNvZhpqRF0Ij/cKUN1MDAAEAAQ==";
            license.LicenseCode = licenseCode;
        }

        public void InitLicense()
        {
            if (license.Status != LicenseStatus.Valid)
            {
                throw new Exception(license.GetAllStatusExceptionsAsString());
            }
            else
            {
                License.Evaluation = license.IsEvaluationLicense();
                License.EvaluationExpired = license.IsEvaluationExpired();

                short maxUsageDays = license.MaxUsageDays;
                DateTime dateGenerated = license.DateGenerated;
                License.ExpirationDate = dateGenerated.AddDays(maxUsageDays);

                License.WindowsForms = isFeatureAvailable(FeatureList.Windows.ToInt());
                License.AutoSync = isFeatureAvailable(FeatureList.AutoSync.ToInt());
                // License.ConsoleSync = isFeatureAvailable(FeatureList.ConsoleSync.ToInt());

                License.CompanyName = GetDataField("CompanyName");
                License.CompanyWebSite = GetDataField("CompanyWebSite");
                License.RegisteredToFirstName = GetDataField("FirstName");
                License.RegisteredToLastName = GetDataField("LastName");
                License.RegisteredToPhoneNumber = GetDataField("PhoneNumber");
                License.RegisteredToEmailAddress = GetDataField("EmailAddress");
                License.ComputerName = GetDataField("ComputerName");
                License.LicenseKey = GetDataField("LicenseKey");
                License.LicenseType = GetDataField("LicenseType");
                License.MaintenanceExpirationDate = GetDataField("SupportExpDate");
            }
        }


        private bool isFeatureAvailable(int feature)
        {
            bool isPresent = false;

            LicenseFeatures features = license.Features;
            if (features != LicenseFeatures.DoesNotMatter)
            {
                isPresent = license.IsFeaturePresent((LicenseFeatures)feature, false);
            }
            return isPresent;
        }

        private string GetDataField(string fieldName)
        {
            Hashtable userData = license.ParseUserData("#");
            string result = userData[fieldName].ToString();

            return result;
        }

    }
}
