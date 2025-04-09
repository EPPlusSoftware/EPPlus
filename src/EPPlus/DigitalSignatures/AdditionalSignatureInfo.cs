using OfficeOpenXml.DigitalSignatures.XAdES;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.DigitalSignatures
{
    /// <summary>
    /// Details about the signer of a digital signature, such as role, title, address etc.
    /// </summary>
    public class AdditionalSignatureInfo
    {
        internal AdditionalSignatureInfo()
        {
        }

        /// <summary>
        /// Role or Title
        /// </summary>
        public string SignerRoleTitle { get; set; } = null;

        /// <summary>
        /// Address
        /// </summary>
        public string Address1 { get; set; } = null;
        /// <summary>
        /// Address 2
        /// </summary>
        public string Address2 { get; set; } = null;

        /// <summary>
        /// Zip or Postal Code
        /// </summary>
        public string ZipOrPostalCode { get; set; } = null;
        /// <summary>
        /// City
        /// </summary>
        public string City { get; set; } = null;
        /// <summary>
        /// Country or region
        /// </summary>
        public string CountryOrRegion { get; set; } = null;
        /// <summary>
        /// State or province
        /// </summary>
        public string StateOrProvince { get; set; } = null;
    }
}
