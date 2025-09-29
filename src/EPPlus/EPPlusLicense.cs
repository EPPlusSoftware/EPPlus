#if(Core)
using Microsoft.Extensions.Configuration;
#endif
using OfficeOpenXml.Configuration;
using OfficeOpenXml.Utils;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Globalization;
using System.Linq;

namespace OfficeOpenXml
{
    /// <summary>
    /// Represents a class to set the license 
    /// </summary>
    public class EPPlusLicense
    {
        private static ExcelPackageConfiguration _configuration = new ExcelPackageConfiguration();
        static bool _licenseSet = false;
        internal const string _versionDate = "2025-08-15";
        /// <summary>
        /// The license key used for a commercial license.
        /// </summary>
        public string LicenseKey { get; private set; }
        /// <summary>
        /// The name used for a commercial organization
        /// </summary>
        public string LegalName
        {
            get;
            private set;
        }
        /// <summary>
        /// The source where the license was set.
        /// </summary>
        public EPPlusLicenseSource? Source { get; private set; }
        /// <summary>
        /// The type of license used.
        /// </summary>
        public EPPlusLicenseType? LicenseType { get; private set; }
        /// <summary>
        /// License information from the license key. If no license key has been set, this property contains null.
        /// </summary>
        public EPPlusLicenseInfo LicenseInfo { get; internal set; }
        /// <summary>
        /// If your subscription has expired past the <see cref="EPPlusLicenseInfo.LicenseValidTo"/> date, you can set this flag to get a 15 additional days to renew the license.
        /// </summary>
        public bool ExtendUnderRenewal { get; set; }

        internal bool IsLicenseSet(List<ExcelInitializationError> initErrors)
        {
            if (_licenseSet == true)
            {
                return true;
            }
            else
            {
                _licenseSet = SetLicenseFromConfig(initErrors);
                return _licenseSet;
            }
        }

#if (Core)
        internal void SetConfiguration(IConfiguration configuration)
        {
            // we only set this if it wasn't already set
            if(_configuration.Configuration == null)
            {
                _configuration.Configuration = configuration;
            }
        }
#endif

        /// <summary>
        /// Use this license if you use EPPlus for personal noncommercial usage.
        /// Using this option will tag all created document with the Polyform Noncommercial license.
        /// See https://polyformproject.org/licenses/noncommercial/1.0.0/        
        /// </summary>
        /// <param name="fullName">Your name. This name will go into the Office Properties</param>
        public void SetNonCommercialPersonal(string fullName)
        {
            if (ValidateName(fullName, out string msg) == false)
            {
                throw new LicenseInformationException(msg);
            }
            LegalName = fullName;
            LicenseType = EPPlusLicenseType.NonCommercialPersonal;
            Source = EPPlusLicenseSource.Code;
            LicenseInfo = null;
            LicenseKey = null;
            _licenseSet = true;
        }
        /// <summary>
        /// Use this option if you use EPPlus within a noncommercial organization.
        /// Using this option will tag all created document with the Polyform Noncommercial license. 
        /// <see href="https://polyformproject.org/licenses/noncommercial/1.0.0/"/>
        /// </summary>
        /// <param name="organizationName">The noncommercial organizations name</param>
        public void SetNonCommercialOrganization(string organizationName)
        {
            if (ValidateName(organizationName, out string msg) == false)
            {
                throw new LicenseInformationException(msg);
            }
            LegalName = organizationName;
            LicenseType = EPPlusLicenseType.NonCommercialOrganization;
            Source = EPPlusLicenseSource.Code;
            LicenseKey = null;
            LicenseInfo = null;
            _licenseSet = true;
        }

        private bool ValidateName(string name, out string msg)
        {
            if (name.Length < 3)
            {
                msg = "The license holder name must be at least 3 characters.";
                return false;
            }
            if (name.Any(c => c == '/' || c == '\\' || c == '*' || char.IsControl(c) || c == '\t'))
            {
                msg = "The license holder name contains invalid characters";
                return false;
            }
            if (name.Count(x => char.IsLetter(x)) < 2)
            {
                msg = "The license holder name contains to few letters.";
                return false;
            }
            msg = "";
            return true;
        }

        /// <summary>
        /// If you use EPPlus within a commercial organization or for commercial purposes.
        /// This requires a license for EPPlus that can be purchased at https://epplussoftware.com
        /// </summary>
        /// <param name="licenseKey">The license key you received with your license.</param>
        /// <seealso href="https://epplussoftware.com/en/Home/GettingStartedLicenseKey"/>
        public void SetCommercial(string licenseKey)
        {
            _licenseSet = LicenseHandler.ValidateLicenseKey(licenseKey, ExtendUnderRenewal, out EPPlusLicenseInfo licenseInfo, out string msg);
            if (ExtendUnderRenewal && DateTime.UtcNow < licenseInfo.LicenseValidTo.AddDays(-20))
            {
                throw new LicenseNotValidException("ExcelPackage.License.ExtendUnderRenewal should be set to false. It should only have a true value during renewals.");
            }
            LicenseInfo = licenseInfo;
            LicenseKey = licenseKey;
            LicenseType = EPPlusLicenseType.Commercial;
            Source = EPPlusLicenseSource.Code;
            if (_licenseSet == false)
            {
                if (licenseInfo.Status == EPPlusLicenseStatus.InvalidLicenseKey)
                {
                    throw new InvalidLicenseKeyException(msg);
                }
                else
                {
                    throw new LicenseNotValidException(msg);
                }
            }
        }

        internal bool SetLicenseFromConfig(List<ExcelInitializationError> initErrors)
        {
            var v = GetConfigValue("License", initErrors, out bool inEnvironment);

            if (string.IsNullOrEmpty(v))
            {
                ExcelPackage.License.Source = null;
                return false;
            }
            else
            {
                v = v.Trim();
                var s = v.Split([':', ',']);
                if (s[0].Equals("noncommercialorganization", StringComparison.OrdinalIgnoreCase))
                {
                    if (s.Length == 1)
                    {
                        throw new LicenseInformationException("Please specify a name for the noncommercial organization in the app config file. Format noncommercialorganization:[name of your organization]");
                    }
                    v = v.Substring(v.IndexOfAny([':', ',']) + 1);
                    SetNonCommercialOrganization(v.Trim());
                    ExcelPackage.License.Source = inEnvironment ? EPPlusLicenseSource.EnvironmentVariable : EPPlusLicenseSource.ConfigFile;
                    return true;
                }

                if (s[0].Equals("noncommercialpersonal", StringComparison.OrdinalIgnoreCase))
                {
                    if (s.Length == 1)
                    {
                        throw new LicenseInformationException("Please specify your name to be used with the license in the app config file. Format noncommercialpersonal:[your name]");
                    }
                    v = v.Substring(v.IndexOfAny([':', ',']) + 1);
                    SetNonCommercialPersonal(v.Trim());
                    ExcelPackage.License.Source = inEnvironment ? EPPlusLicenseSource.EnvironmentVariable : EPPlusLicenseSource.ConfigFile;
                    return true;
                }
                else
                {
                    if (s.Length > 1 && v.StartsWith("commercial", StringComparison.OrdinalIgnoreCase))
                    {
                        v = v.Substring(v.IndexOfAny([':', ',']) + 1);
                    }
                    SetCommercial(v.Trim());
                    ExcelPackage.License.Source = inEnvironment ? EPPlusLicenseSource.EnvironmentVariable : EPPlusLicenseSource.ConfigFile;
                    return true;
                }
            }
        }
        private static string GetConfigValue(string key, List<ExcelInitializationError> initErrors, out bool inEnvironment)
        {
            var v = ExcelConfigurationReader.GetEnvironmentVariable("EPPlus" + key, EnvironmentVariableTarget.Process, _configuration, initErrors);
            if (string.IsNullOrEmpty(v))
            {
                v = ExcelConfigurationReader.GetEnvironmentVariable("EPPlus" + key, EnvironmentVariableTarget.User, _configuration, initErrors);
                if (string.IsNullOrEmpty(v))
                {
                    v = ExcelConfigurationReader.GetEnvironmentVariable("EPPlus" + key, EnvironmentVariableTarget.Machine, _configuration, initErrors);
                }
            }
            if (string.IsNullOrEmpty(v))
            {
#if (Core)
                v = ExcelConfigurationReader.GetJsonConfigValue($"EPPlus:ExcelPackage:{key}", _configuration, initErrors);

#else
                    v = ExcelConfigurationReader.GetValueFromAppSettings($"EPPlus:ExcelPackage:{key}", _configuration, initErrors);
                    if(string.IsNullOrEmpty(v))
                    {
                        v = ExcelConfigurationReader.GetValueFromAppSettings($"EPPlus:ExcelPackage.{key}", _configuration, initErrors);
                    }
#endif
                inEnvironment = false;
            }
            else
            {
                inEnvironment = true;
            }
            return v;
        }
        /// <summary>
        /// Removes the active license.
        /// </summary>
        public void RemoveActiveLicense()
        {
            LicenseType = null;
            LicenseKey = null;
            Source = null;
            LegalName = null;
            LicenseInfo = null;
            _licenseSet = false;
        }
    }
}
