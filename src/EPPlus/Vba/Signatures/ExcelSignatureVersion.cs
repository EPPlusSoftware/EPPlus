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
using System.Security.Cryptography.X509Certificates;
using System.Security.Cryptography.Pkcs;
using OfficeOpenXml.VBA.Signatures;
using OfficeOpenXml.VBA;

namespace OfficeOpenXml.Vba.Signatures
{
    /// <summary>
    /// Signature version settings.
    /// </summary>
    public class ExcelSignatureVersion
    {
        internal ExcelSignatureVersion(EPPlusVbaSignature signature, VbaSignatureHashAlgorithm hashAlgorithm)
        {
            SignatureHandler = signature;
            CreateSignatureOnSave = SignatureHandler.ReadSignature();
            Part = SignatureHandler.Part;
            Verifier = SignatureHandler.Verifier;
            HashAlgorithm = hashAlgorithm;
        }
        /// <summary>
        /// A boolean indicating if a signature for the VBA project will be created when the package is saved.
        /// Default is true
        /// </summary>
        public bool CreateSignatureOnSave { get; set; } = true;
        /// <summary>
        /// The verifyer
        /// </summary>
        public SignedCms Verifier { get; internal set; }
        /// <summary>
        /// The hash algorithm used.
        /// </summary>
        public VbaSignatureHashAlgorithm HashAlgorithm
        {
            get;
            set;
        }
        internal Packaging.ZipPackagePart Part { get; set; }
        internal readonly EPPlusVbaSignature SignatureHandler;
        internal X509Certificate2 Certificate
        {
            get
            {
                return SignatureHandler.Certificate;
            }
            set
            {
                SignatureHandler.Certificate = value;
            }
        }
        internal void CreateSignature(ExcelVbaProject project)
        {
            SignatureHandler.Context.HashAlgorithm = HashAlgorithm;
            SignatureHandler.CreateSignature(project);
        }
    }
}
