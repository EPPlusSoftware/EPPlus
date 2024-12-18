using System;
using System.Security.Cryptography.Xml;
using OfficeOpenXml.Constants;
using OfficeOpenXml.Packaging;

namespace OfficeOpenXml.DigitalSignatures
{
    internal class SignatureProject
    {
        internal const string PartUri = @"/_xmlsignatures/sig1.xml";
        internal ZipPackagePart part;

        internal SignatureProject(ExcelWorkbook wb) 
        {
            part = wb._package.ZipPackage.CreatePart(new Uri(PartUri), ContentTypes.xmlSignatures);
        }
    }
}
