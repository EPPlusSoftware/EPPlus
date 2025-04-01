using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.DigitalSignatures
{
    /// <summary>
    /// Digital Signature Types
    /// </summary>
    public enum DigitalSignatureType
    {
        /// <summary>
        /// No visible representation of the digital signature
        /// </summary>
        Invisible = 1,
        /// <summary>
        /// There is a visible representation of the digital signature
        /// </summary>
        SignatureLine = 2,
    }
}
