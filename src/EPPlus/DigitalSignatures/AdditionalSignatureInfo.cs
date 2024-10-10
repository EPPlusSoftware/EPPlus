using OfficeOpenXml.DigitalSignatures.XAdES;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.DigitalSignatures
{
    public class AdditionalSignatureInfo
    {
        internal AdditionalSignatureInfo()
        {
        }

        public string SignerRoleTitle { get; set; } = null;

        public string Address1 { get; set; } = null;
        public string Address2 { get; set; } = null;

        public string ZIPorPostalCode { get; set; } = null;
        public string City { get; set; } = null;
        public string CountryOrRegion { get; set; } = null;
        public string StateOrProvince { get; set; } = null;
    }
}
