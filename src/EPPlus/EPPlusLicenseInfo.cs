using System;

namespace OfficeOpenXml
{
    public class EPPlusLicenseInfo
    {
        internal EPPlusLicenseInfo()
        {
            
        }
        public string LicenseNumber { get; internal set; }
        public byte LicenseType { get; internal set; }
        public DateTime LicenseValidFrom { get; internal set; }
        public DateTime LicenseValidTo { get; internal set; }
        public int NumberOfLicenses { get; internal set; }
    }
}
