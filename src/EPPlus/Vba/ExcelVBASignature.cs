/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  01/27/2020         EPPlus Software AB       Initial release EPPlus 5
 *************************************************************************************************/
using System;
using System.Linq;
using System.Security.Cryptography.X509Certificates;
using OfficeOpenXml.Utils;
using OfficeOpenXml.Utils.CompundDocument;
using System.Security.Cryptography.Pkcs;
using OfficeOpenXml.Packaging;
using OfficeOpenXml.VBA.Signatures;
using OfficeOpenXml.Vba.Signatures;

namespace OfficeOpenXml.VBA
{
    /// <summary>
    /// The VBA project's code signature properties
    /// </summary>
    public class ExcelVbaSignature
    {
        internal ExcelVbaSignature(Packaging.ZipPackagePart vbaPart)
        {
            _vbaPart = vbaPart;
            LegacySignature = new ExcelSignatureVersion(new EPPlusVbaSignatureLegacy(vbaPart), VbaSignatureHashAlgorithm.MD5);
            AgileSignature = new ExcelSignatureVersion(new EPPlusVbaSignatureAgile(vbaPart), VbaSignatureHashAlgorithm.SHA1);
            V3Signature = new ExcelSignatureVersion(new EPPlusVbaSignatureV3(vbaPart), VbaSignatureHashAlgorithm.SHA1);
            
            if (LegacySignature.Certificate != null) _certificate = LegacySignature.Certificate;
            if (_certificate == null && AgileSignature.Certificate != null) _certificate = AgileSignature.Certificate;
            if (_certificate == null && V3Signature.Certificate != null) _certificate = V3Signature.Certificate;
        }

        internal readonly ZipPackagePart _vbaPart = null;
        private X509Certificate2 _certificate;

        /// <summary>
        /// The certificate to sign the VBA project.
        /// <remarks>
        /// This certificate must have a private key.
        /// There is no validation that the certificate is valid for codesigning, so make sure it's valid to sign Excel files (Excel 2010 is more strict that prior versions).
        /// </remarks>
        /// </summary>
        public X509Certificate2 Certificate 
        { 
            get
            {
                return _certificate;
            }
            set 
            {
                _certificate = value;
            } 
        }
        /// <summary>
        /// The verifier (legacy format)
        /// </summary>
        public SignedCms Verifier { get; internal set; }
        internal CompoundDocument Signature { get; set; }
        internal Packaging.ZipPackagePart Part { get; set; }
        internal void Save(ExcelVbaProject proj)
        {
            if (Certificate == null) return;
            
            //Legacy signature
            if (LegacySignature.CreateSignatureOnSave)
            {
                LegacySignature.SignatureHandler.Certificate = Certificate;
                LegacySignature.CreateSignature(proj);
            }
            else if(Part?.Uri != null && Part.Package.PartExists(Part.Uri))
            {                
                Part.Package.DeletePart(Part.Uri);
            }

            //Agile signature
            var p = AgileSignature.Part;
            if (AgileSignature.CreateSignatureOnSave)
            {
                AgileSignature.SignatureHandler.Certificate = Certificate;
                AgileSignature.CreateSignature(proj);
            }
            else if (p?.Uri != null && p.Package.PartExists(p.Uri))
            {
                p.Package.DeletePart(p.Uri);
            }

            //V3 signature
            p = V3Signature.Part;
            if (V3Signature.CreateSignatureOnSave)
            {
                V3Signature.Certificate = Certificate;
                V3Signature.CreateSignature(proj);
            }
            else if (p?.Uri != null && p.Package.PartExists(p.Uri))
            {
                p.Package.DeletePart(p.Uri);
            }
        }
        /// <summary>
        /// Settings for the legacy signing.
        /// </summary>
        public ExcelSignatureVersion LegacySignature { get; set; }
        /// <summary>
        /// Settings for the agile vba signing. 
        /// The agile signature adds a hash that is calculated for user forms data in the vba project (designer streams). 
        /// </summary>
        public ExcelSignatureVersion AgileSignature { get; set; }
        /// <summary>
        /// Settings for the V3 vba signing.
        /// The V3 signature includes more coverage for data in the dir and project stream in the hash, not covered by the legacy and agile signatures.
        /// </summary>
        public ExcelSignatureVersion V3Signature { get; set; }
    }
}
