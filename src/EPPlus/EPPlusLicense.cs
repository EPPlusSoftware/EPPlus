using OfficeOpenXml.Drawing.Chart.ChartEx;
using OfficeOpenXml.Table.PivotTable.Calculation.ShowDataAs;
using System;
using System.Globalization;
using System.Reflection;

namespace OfficeOpenXml
{
    /// <summary>
    /// Represents a class to set the license 
    /// </summary>
    public class EPPlusLicense
    {
        internal const string _versionDate = "2024-10-01";
        internal string LicenseKey { get; set; }
        public string LegalName { get; private set; }
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
    }
    /// <summary>
    /// The type of license used.
    /// </summary>
    public enum EPPlusLicenseType
    {
        /// <summary>
        /// Use this license if you use EPPlus for personal non-commercial usage.
        /// </summary>
        NonCommercialPersonal = 0,
        /// <summary>
        /// Use this license if you use EPPlus representing a non-commercial organization.
        /// </summary>
        NonCommercialOrganization = 1,
        /// <summary>
        /// If you use EPPlus within a commercial organization or for commercial purposes.
        /// </summary>
        Commercial = 2,
    }
}
