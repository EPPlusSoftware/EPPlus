/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  09/05/2022         EPPlus Software AB       EPPlus 6.1
 *************************************************************************************************/
using OfficeOpenXml.Packaging;
using OfficeOpenXml.Utils;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Cryptography.Pkcs;
using System.Security.Cryptography.X509Certificates;
using System.Text;

namespace OfficeOpenXml.VBA.Signatures
{
    internal abstract class EPPlusVbaSignature
    {
        public EPPlusVbaSignature(ZipPackagePart vbaPart, ExcelVbaSignatureType signatureType)
        {
            _vbaPart = vbaPart;
            _signatureType = signatureType;
        }

        private readonly ZipPackagePart _vbaPart;
        private readonly ExcelVbaSignatureType _signatureType;
        internal ZipPackagePart Part
        {
            get;
            set;
        }

        private string SchemaRelation
        {
            get
            {
                switch(_signatureType)
                {
                    case ExcelVbaSignatureType.Legacy:
                        return VbaSchemaRelations.Legacy;
                    case ExcelVbaSignatureType.Agile:
                        return VbaSchemaRelations.Agile;
                    case ExcelVbaSignatureType.V3:
                        return VbaSchemaRelations.V3;
                    default:
                        return VbaSchemaRelations.Legacy;
                }
            }
        }

        public X509Certificate2 Certificate { get; set; }
        public SignedCms Verifier { get; internal set; }

        public EPPlusSignatureContext Context { get; set; }

        internal void ReadSignature()
        {

            if (_vbaPart == null) return;
            var rel = _vbaPart.GetRelationshipsByType(SchemaRelation).FirstOrDefault();
            if(rel != null)
            {
                var uri = UriHelper.ResolvePartUri(rel.SourceUri, rel.TargetUri);
                Part = _vbaPart.Package.GetPart(uri);
                Context = new EPPlusSignatureContext(_signatureType);
                var signature = SignatureReader.ReadSignature(Part, _signatureType, Context);
                Certificate = signature.Certificate;
                Verifier = signature.Verifier;
            }
            else
            {
                Certificate = null;
                Verifier = null;
            }
        }

        internal void CreateSignature(ExcelVbaProject project)
        {
            byte[] certStore = CertUtil.GetSerializedCertStore(Certificate.RawData);
            if (Certificate == null)
            {
                SignaturePartUtil.DeleteParts(Part);
                return;
            }

            if (Certificate.HasPrivateKey == false)    //No signature. Remove any Signature part
            {
                var storeCert = CertUtil.GetCertificate(Certificate.Thumbprint);
                if (storeCert != null)
                {
                    Certificate = storeCert;
                }
                else
                {
                    SignaturePartUtil.DeleteParts(Part);
                    return;
                }
            }
            using (var ms = RecyclableMemory.GetStream())
            {
                var bw = new BinaryWriter(ms);
                Verifier = CertUtil.SignProject(project, this, Context);
                var cert = Verifier.Encode();
                var signatureBytes = CertUtil.CreateBinarySignature(ms, bw, certStore, cert);
                Part = SignaturePartUtil.GetPart(project, this, SchemaRelation);
                Part.GetStream(FileMode.Create).Write(signatureBytes, 0, signatureBytes.Length);
            }
            
        }
    }
}
