namespace OfficeOpenXml
{
    /// <summary>
    /// Represents a class to set the license 
    /// </summary>
    public class EPPlusLicense
    {
        string _licenseKey;
        public LicenseType LicenseType { get; set; }
        /// <summary>
        /// Use this license if you use EPPlus for personal non-commercial usage.
        /// Using this option will tag all created document with the Polyform Non-Commercial license.
        /// See https://polyformproject.org/licenses/noncommercial/1.0.0/        
        /// </summary>
        /// <param name="fullName">Your name. This name will go into the Office Properties</param>
        public void SetLicenseNonCommercialPersonal(string fullName)
        {

        }
        /// <summary>
        /// User this option if you use EPPlus within a non-commercial organization.
        /// Using this option will tag all created document with the Polyform Non-Commercial license. 
        /// See https://polyformproject.org/licenses/noncommercial/1.0.0/        
        /// </summary>
        /// <param name="organizationName">The non-commercial organziations name</param>
        public void SetLicenseNonCommercialOrganization(string organizationName)
        {

        }
        /// <summary>
        /// If you use EPPlus within a commercial organization or for commercial purposes.
        /// This requires a license for EPPlus that can be purchased at https://epplussoftware.com
        /// </summary>
        /// <param name="licenseKey">The licens key you recieved with your license</param>
        public void SetLicenseCommercial(string licenseKey)
        {

        }
    }
    /// <summary>
    /// The type of license used.
    /// </summary>
    public enum LicenseType
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
        Commercial = 2
    }
}
