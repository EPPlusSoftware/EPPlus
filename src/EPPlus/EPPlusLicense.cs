using OfficeOpenXml.Configuration;
using OfficeOpenXml.Utils;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;

namespace OfficeOpenXml
{
    /// <summary>
    /// Represents a class to set the license 
    /// </summary>
    public class EPPlusLicense
    {
        //ExcelPackage _pck;
        private static ExcelPackageConfiguration _configuration = new ExcelPackageConfiguration();
        //internal EPPlusLicense(ExcelPackage pck)
        //{
        //    _pck = pck;
        //}
        internal const string _versionDate = "2024-10-01";
        /// <summary>
        /// The license key used for a commercial license.
        /// </summary>
        public string LicenseKey { get; private set; }
        private string _legalName="";
        /// <summary>
        /// The name used for a commercial organization
        /// </summary>
        public string LegalName 
        {
            get
            {
                return _legalName;
            }
            private set
            {
                if(value==null || value.RemoveAllWhiteSpaces().Length<4)
                {
                    throw new LicenseException("License: Legal name must contain more than 3 non-whitespace characters.");
                }
                _legalName = value.Trim();
            }
        }
        /// <summary>
        /// The type of license used.
        /// </summary>
        public EPPlusLicenseType? LicenseType { get; private set; }
        /// <summary>
        /// License information from the license key. If no license key has been set, this propery contains null;
        /// </summary>
        public EPPlusLicenseInfo LicenseInfo { get; internal set; }

        /// <summary>
        /// Use this license if you use EPPlus for personal non-commercial usage.
        /// Using this option will tag all created document with the Polyform Non-Commercial license.
        /// See https://polyformproject.org/licenses/noncommercial/1.0.0/        
        /// </summary>
        /// <param name="fullName">Your name. This name will go into the Office Properties</param>
        public void SetLicenseNonCommercialPersonal(string fullName)
        {
            LegalName = fullName;
            LicenseType = EPPlusLicenseType.NonCommercialPersonal;
            LicenseInfo = null;
        }
        /// <summary>
        /// User this option if you use EPPlus within a non-commercial organization.
        /// Using this option will tag all created document with the Polyform Non-Commercial license. 
        /// See https://polyformproject.org/licenses/noncommercial/1.0.0/        
        /// </summary>
        /// <param name="organizationName">The non-commercial organziations name</param>
        public void SetLicenseNonCommercialOrganization(string organizationName)
        {
            LegalName = organizationName;
            LicenseType = EPPlusLicenseType.NonCommercialOrganization;
            LicenseInfo = null;
        }
        /// <summary>
        /// If you use EPPlus within a commercial organization or for commercial purposes.
        /// This requires a license for EPPlus that can be purchased at https://epplussoftware.com
        /// </summary>
        /// <param name="licenseKey">The licens _key you recieved with your license</param>
        public void SetLicenseCommercial(string licenseKey)
        {
            LicenseInfo = new EPPlusLicenseInfo();
            if (LicenseHandler.ValidateLicenseKey(licenseKey, LicenseInfo))
            {
                LicenseKey = licenseKey;
                LicenseType = EPPlusLicenseType.Commercial;
                if(LicenseInfo.LicenseValidFrom > DateTime.Today)
                {
                    throw new LicenseException($"This license is not valid until {LicenseInfo.LicenseValidFrom:d}.");
                }
                var vd = DateTime.Parse(_versionDate, CultureInfo.InvariantCulture);
                if (LicenseInfo.LicenseValidTo < vd)
                {
                    throw new LicenseHasExpiredException($"This license expired {LicenseInfo.LicenseValidTo:d} and is not valid for this version of EPPlus({_versionDate:d}).");
                }
            }
        }

        internal bool SetLicenseFromConfig(List<ExcelInitializationError> initErrors)
        {
            if (Debugger.IsAttached == false)   //This check is only performed if a debugger is attached. 
            {
                return true;
            }
            var v = GetConfigValue("License", initErrors, out bool inEnvironment);

            if (string.IsNullOrEmpty(v))
            {
                inEnvironment = false;
                return false;
            }
            else
            {
                v = v.Trim();
                var s = v.Split([':',',']);
                if (s[0].Equals("noncommercialorganization", StringComparison.OrdinalIgnoreCase))
                {
                    if (s.Length == 1)
                    {
                        throw new LicenseException("Please specify a name for the non-commercial organization in the app config file. Format noncommercialorganization:[name of your organization]");
                    }
                    v = v.Substring(v.IndexOfAny([':', ',']) + 1);
                    SetLicenseNonCommercialOrganization(v.Trim());
                    return true;
                }

                if (s[0].Equals("noncommercialpersonal", StringComparison.OrdinalIgnoreCase))
                {
                    if (s.Length == 1)
                    {
                        throw new LicenseException("Please specify your name to be used with the license in the app config file. Format noncommercialpersonal:[your name]");
                    }
                    v = v.Substring(v.IndexOfAny([':', ',']) + 1);
                    SetLicenseNonCommercialPersonal(v.Trim());
                    return true;
                }
                else
                {
                    if (s.Length > 1 && v.StartsWith("commercial", StringComparison.OrdinalIgnoreCase))
                    {
                        v = v.Substring(v.IndexOfAny([':',',']) + 1);
                    }
                    SetLicenseCommercial(v.Trim());
                    return true;
                }
            }
        }
        private static string GetConfigValue(string key, List<ExcelInitializationError> initErrors, out bool inEnvironment)
        {
            var v = ExcelConfigurationReader.GetEnvironmentVariable("EPPlus" + key, EnvironmentVariableTarget.User, _configuration, initErrors);
            if (string.IsNullOrEmpty(v))
            {
                v = ExcelConfigurationReader.GetEnvironmentVariable("EPPlus" + key, EnvironmentVariableTarget.Process, _configuration, initErrors);
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

    }
    /// <summary>
    /// The type of license used.
    /// </summary>
    public enum EPPlusLicenseType
    {
        /// <summary>
        /// If you use EPPlus within a commercial organization or for commercial purposes.
        /// </summary>
        Commercial = 0,
        /// <summary>
        /// Use this license if you use EPPlus for personal non-commercial usage.
        /// </summary>
        NonCommercialPersonal = 1,
        /// <summary>
        /// Use this license if you use EPPlus representing a non-commercial organization.
        /// </summary>
        NonCommercialOrganization = 2
    }
}
