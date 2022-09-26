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

namespace OfficeOpenXml.VBA
{
    /// <summary>
    /// The code signature properties of the project
    /// </summary>
    public class ExcelVbaSignature
    {
        internal ExcelVbaSignature(Packaging.ZipPackagePart vbaPart)
        {
            _vbaPart = vbaPart;
            _legacySignature = new EPPlusVbaSignatureLegacy(vbaPart);
            _agileSignature = new EPPlusVbaSignatureAgile(vbaPart);
            _v3Signature = new EPPlusVbaSignatureV3(vbaPart);
            ReadSignatures();
        }

        private readonly ZipPackagePart _vbaPart = null;
        private readonly EPPlusVbaSignature _legacySignature;
        private readonly EPPlusVbaSignature _agileSignature;
        private readonly EPPlusVbaSignature _v3Signature;
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
                _legacySignature.Certificate = value;
                if(_agileSignature != null) _agileSignature.Certificate = value;
                if(_v3Signature != null) _v3Signature.Certificate = value;
            } 
        }
        /// <summary>
        /// The verifier (legacy format)
        /// </summary>
        public SignedCms Verifier { get; internal set; }
        /// <summary>
        /// The verifier (agile format)
        /// </summary>
        public SignedCms VerifierAgile { get; internal set; }
        /// <summary>
        /// The verifier (V3 format)
        /// </summary>
        public SignedCms VerifierV3 { get; internal set; }
        internal CompoundDocument Signature { get; set; }
        internal Packaging.ZipPackagePart Part { get; set; }
        internal Packaging.ZipPackagePart PartAgile { get; set; }
        internal Packaging.ZipPackagePart PartV3 { get; set; }
        private void ReadSignatures()
        {
            if (_vbaPart == null) return;

            // Legacy signature
            if (_legacySignature != null)
            {
                _legacySignature.ReadSignature();
                Part = _legacySignature.Part;
                Certificate = _legacySignature.Certificate;
                Verifier = _legacySignature.Verifier;
            }

            if (_agileSignature != null)
            {
                // Agile signature
                _agileSignature.ReadSignature();
                PartAgile = _agileSignature.Part;
                if (Certificate == null)
                {
                    Certificate = _agileSignature.Certificate;
                }
                VerifierAgile = _agileSignature.Verifier;
            }
            if (_v3Signature != null)
            {
                // V3 signature
                _v3Signature.ReadSignature();
                PartV3 = _v3Signature.Part;
                if (Certificate == null)
                {
                    Certificate = _v3Signature.Certificate;
                }
                VerifierV3 = _v3Signature.Verifier;
            }
        }

        internal void Save(ExcelVbaProject proj)
        {
            if (Certificate == null) return;
            
            //Legacy signature
            if (CreateLegacySignatureOnSave)
            {
                _legacySignature.Certificate = Certificate;
                _legacySignature.CreateSignature(proj);
            }
            else if(Part?.Uri != null && Part.Package.PartExists(Part.Uri))
            {                
                Part.Package.DeletePart(Part.Uri);
            }

            //Agile signature
            if (CreateAgileSignatureOnSave)
            {
                _agileSignature.Certificate = Certificate;
                _agileSignature.CreateSignature(proj);
            }
            else if (PartAgile?.Uri != null && PartAgile.Package.PartExists(PartAgile.Uri))
            {
                PartAgile.Package.DeletePart(PartAgile.Uri);
            }

            //V3 signature
            if (CreateV3SignatureOnSave)
            {
                _v3Signature.Certificate = Certificate;
                _v3Signature.CreateSignature(proj);
            }
            else if (PartV3?.Uri != null && PartV3.Package.PartExists(PartV3.Uri))
            {
                PartV3.Package.DeletePart(PartV3.Uri);
            }
        }
        /// <summary>
        /// A boolean indicating if a legacy signature for the VBA project should be created when the package is saved.
        /// Default is true
        /// </summary>
        public bool CreateLegacySignatureOnSave { get; set; } = true;
        /// <summary>
        /// A boolean indicating if a agile signature for the VBA project should be created when the package is saved.
        /// Default is true
        /// /// </summary>
        public bool CreateAgileSignatureOnSave { get; set; } = true;
        /// <summary>
        /// A boolean indicating if a v3 signature for the VBA project should be created when the package is saved.
        /// Default is true
        /// </summary>
        public bool CreateV3SignatureOnSave { get; set; } = true;
    }
}
