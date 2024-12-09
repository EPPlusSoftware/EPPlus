using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Drawing
{
    /// <summary>
    /// Signature line types.
    /// </summary>
    public enum eSignatureLineType
    {
        /// <summary>
        /// Signature line stamp which contains a SignatureImage as the signature.
        /// </summary>
        Stamp,
        /// <summary>
        /// Signature line which can contain either text or an image as the signature.
        /// </summary>
        SignatureLine
    }
}
