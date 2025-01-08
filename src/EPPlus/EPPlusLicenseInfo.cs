using System;

namespace OfficeOpenXml
{
    /// <summary>
    /// License information about a commercial license.
    /// </summary>
    public partial class EPPlusLicenseInfo
    {
        internal EPPlusLicenseInfo()
        {
            
        }
        /// <summary>
        /// The license number
        /// </summary>
        public string LicenseNumber { get; internal set; }
        /// <summary>
        /// The type of license
        /// </summary>
        public EPPlusCommercialLicenseType LicenseType { get; internal set; }
        /// <summary>
        /// The license valid from date.
        /// </summary>
        public DateTime LicenseValidFrom { get; internal set; }
        /// <summary>
        /// The license valid to date.
        /// For subsciption licenses, the license will work for 30 days after this date to allow the licens to be renewed.
        /// For perpetual licenses, you will not be able to update to major/minor version of EPPlus release after this date.
        /// </summary>
        public DateTime LicenseValidTo { get; internal set; }
        /// <summary>
        /// The number of developers covered by this license.
        /// </summary>
        public int NumberOfLicensedDevelopers { get; internal set; }        
        /// <summary>
        /// 
        /// </summary>
        public EPPlusLicenseStatus Status
        {
            get;
            internal set;
        }
    }
}
