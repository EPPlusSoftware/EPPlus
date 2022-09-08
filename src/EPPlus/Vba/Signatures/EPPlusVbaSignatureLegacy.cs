using OfficeOpenXml.Packaging;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.VBA.Signatures
{
    internal class EPPlusVbaSignatureLegacy : EPPlusVbaSignature
    {
        public EPPlusVbaSignatureLegacy(ZipPackagePart part) 
            : base(part, ExcelVbaSignatureType.Legacy)
        {
        }
    }
}
