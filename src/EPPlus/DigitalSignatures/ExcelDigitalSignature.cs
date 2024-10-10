using System;
using System.Security.Cryptography.X509Certificates;
using OfficeOpenXml.Constants;
using OfficeOpenXml.Packaging;
using System.Security.Cryptography;
using System.Security.Cryptography.Xml;
using System.Xml;
using OfficeOpenXml.DigitalSignatures.XAdES;
using System.Collections.Generic;
using OfficeOpenXml.VBA;
using System.Linq;

namespace OfficeOpenXml.DigitalSignatures
{
    internal class ExcelDigitalSignature : XmlHelper
    {
        internal ZipPackagePart _part;
        ZipPackagePart _originPart;
        ExcelWorkbook _wb;

        public X509Certificate2 Certificate { get; set; } = null;

        const string _originPartUri = @"/_xmlsignatures/origin.sigs";
        const string relType = "http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/signature";
        const string relTypeOrigin = "http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/origin";
        internal const string PartUriBase = @"/_xmlsignatures/sig{0}.xml";
        internal string PartUri = "";

        string _digestMethod = DigestMethods.SHA1;
        string _referenceType = "http://www.w3.org/2000/09/xmldsig#Object";

        XmlDocument _doc;

        bool shouldSave = true;
        bool wasRead = false;
        private SignatureProperty signatureProperty;
        public AdditionalSignatureInfo SignerInformation = new AdditionalSignatureInfo();

        public string PurposeForSigning = "";
        public CommitmentType commitmentType = CommitmentType.None;
        /// <summary>
        /// Signature is verified to be valid
        /// </summary>
        public bool Verified { get; private set; } = false;
        QualifyingProperties qualifyingProperties;

        /// <summary>
        /// Image of the signature if signature type is SignatureLine
        /// </summary>
        internal DigitalSignatureLine SignatureLine = null;

        internal ExcelDigitalSignature(ExcelWorkbook wb, XmlNamespaceManager ns, ZipPackagePart part, int num) : base(ns)
        {
            PartUri = string.Format(PartUriBase, num);
            _part = part;

            _wb = wb;
            _doc = new XmlDocument()
            {
                PreserveWhitespace = true,
            };
                //Full read only REALLY relevant for verification of signature
            _doc.Load(part.GetStream());

            var officeObj = _doc.SelectSingleNode("//*[@Id='idOfficeV1Details']");

            if(officeObj != null)
            {
                signatureProperty = new SignatureProperty((XmlElement)officeObj, SignerInformation);
            }

            var signedPropertiesNode = _doc.SelectSingleNode("//*[@Id='idSignedProperties']");
            qualifyingProperties = new QualifyingProperties((XmlElement)signedPropertiesNode, SignerInformation);

            string keyInfo = _doc.GetElementsByTagName("KeyInfo")[0].InnerText;
            string serialInFile = qualifyingProperties.SignedProps.SignatureProps.Serial;

            X509Store store = new X509Store(StoreLocation.CurrentUser);
            store.Open(OpenFlags.ReadOnly);
            foreach (var cert in store.Certificates)
            {
                var bytes = cert.GetSerialNumber();
                bytes = bytes.Reverse().ToArray();
                var serialAsDecimals = SignedSignatureProperites.BytesToNumericString(bytes);
                if (serialAsDecimals == serialInFile)
                {
                    Certificate = cert;
                    break;
                }
            }

            shouldSave = Certificate != null;
            wasRead = true;
        }

        internal ExcelDigitalSignature(ExcelWorkbook wb, XmlNamespaceManager ns, int num) : base(ns)
        {
            _wb = wb;
            _doc = new XmlDocument()
            {
                PreserveWhitespace = true,
            };

            PartUri = string.Format(PartUriBase, num);

            _part = wb._package.ZipPackage.CreatePart(new Uri(PartUri, UriKind.Relative), ContentTypes.xmlSignatures);
            var uri = new Uri(_originPartUri, UriKind.Relative);
            if (!wb._package.ZipPackage.PartExists(uri))
            {
                _originPart = wb._package.ZipPackage.CreatePart(uri, ContentTypes.signatureOrigin, CompressionLevel.Default, "sigs");
                wb._package.ZipPackage.CreateRelationship(_originPartUri, TargetMode.Internal, relTypeOrigin);
                var stream = _originPart.GetStream();
                stream.Write([], 0, 0);
            }
            else
            {
                _originPart = wb._package.ZipPackage.GetPart(uri);
            }
            _originPart.CreateRelationship(string.Format("sig{0}.xml", num), TargetMode.Internal, relType);
        }

        private string GetCommitmentTypeString(CommitmentType type)
        {
            switch (type) 
            {
                case CommitmentType.None:
                    return "None";
                case CommitmentType.Approved:
                    return "Approved this document";
                case CommitmentType.Created:
                    return "Created this document";
                case CommitmentType.CreatedAndApproved:
                    return "Created and approved this document";
                default: 
                    throw new NotImplementedException();
            }
        }

        internal void Save()
        {
            if (shouldSave)
            {
                var signatureComments = new List<string>
                {
                    PurposeForSigning
                };

                qualifyingProperties = new QualifyingProperties
                    ("xd", Certificate, GetCommitmentTypeString(commitmentType), signatureComments, SignerInformation);

                var docTest = qualifyingProperties.GetDocument();
                _doc = docTest;

                RSA key;
#if NET35
                key = (RSA)Certificate.PrivateKey;
#else
                key = Certificate.GetRSAPrivateKey();
#endif
                ExcelSignedXml signedXml = new(_doc)
                {
                    SigningKey = key,
                };

                signedXml.Signature.Id = "idPackageSignature";
                signedXml.SignedInfo.CanonicalizationMethod = SignedXml.XmlDsigCanonicalizationUrl;
                signedXml.SignedInfo.SignatureMethod = SignedXml.XmlDsigRSASHA1Url;

                signedXml.KeyInfo = new KeyInfo();
                signedXml.KeyInfo.AddClause(new KeyInfoX509Data(Certificate));

                CreatePackageReference(ref signedXml);
                CreateOfficeReference(ref signedXml);
                CreatePropertiesReference(ref signedXml);

                var value = signedXml.SignatureValue;

                signedXml.ComputeSignature();

                var value2 = signedXml.SignatureValue;

                XmlElement xmlDigitalSignature = signedXml.GetXml();

                var outPutDoc = new XmlDocument()
                {
                    PreserveWhitespace = true,
                };

                var node = outPutDoc.ImportNode(xmlDigitalSignature, true);
                outPutDoc.AppendChild(node);

                var sigValue = outPutDoc.GetElementsByTagName("SignatureValue")[0];
                sigValue.InnerText = Convert.ToBase64String(signedXml.SignatureValue, Base64FormattingOptions.InsertLineBreaks);

                var doc = new XmlDocument();
                doc.LoadXml(outPutDoc.OuterXml);

                Verified = VerifyXmlFile(doc, key);

                var declaration = outPutDoc.CreateXmlDeclaration("1.0", "UTF-8", "");
                outPutDoc.InsertBefore(declaration, node);

                var stream = _part.GetStream();
                stream.Position = 0;

                outPutDoc.Save(stream);

                if (stream.Length > stream.Position)
                {
                    stream.SetLength(stream.Position);
                }
            }
        }

        /// <summary>
        /// Verify that @doc is a valid signed xml file according to the given key
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="Key"></param>
        /// <returns></returns>
        public static bool VerifyXmlFile(XmlDocument doc, RSA Key)
        {
            // Create a new SignedXml object and pass it
            // the XML document class.
            SignedXml signedXml = new SignedXml(doc);

            // Find the "Signature" node and create a new
            // XmlNodeList object.
            XmlNodeList nodeList = doc.GetElementsByTagName("Signature");

            // Load the signature node.
            signedXml.LoadXml((XmlElement)nodeList[0]);

            // Check the signature and return the result.
            return signedXml.CheckSignature(Key);
        }
        public Reference CreatePackageReference(ref ExcelSignedXml signedXml)
        {
            Reference packageReference = new()
            {
                Type = _referenceType,
                Uri = "#idPackageObject"
            };
            packageReference.DigestMethod = _digestMethod;

            var packageObj = new DataObject();

            DigSigManifest manifest = _wb._package.ZipPackage.Manifest;
            var packageObjProps = new SignatureProperty("#idPackageSignature", "idSignatureTime", DateTime.Now);

            var packageObjectDoc = new XmlDocument();

            var docManifest = manifest.GetDoc();
            var docProps = packageObjProps.GetXMLDocument();

            var rootpackageReference = packageObjectDoc.CreateElement("Object", "http://www.w3.org/2000/09/xmldsig#");
            packageObjectDoc.AppendChild(rootpackageReference);
            var manifestImport = packageObjectDoc.ImportNode(docManifest.DocumentElement, true);
            var propsImport = packageObjectDoc.ImportNode(docProps.DocumentElement, true);

            rootpackageReference.AppendChild(manifestImport);
            rootpackageReference.AppendChild(propsImport);

            packageObj.LoadXml(packageObjectDoc.DocumentElement);
            packageObj.Id = "idPackageObject";

            signedXml.AddObject(packageObj);
            signedXml.AddReference(packageReference);

            return packageReference;
        }

        public Reference CreateOfficeReference(ref ExcelSignedXml signedXml)
        {
            Reference officeReference = new()
            {
                Type = _referenceType,
                Uri = "#idOfficeObject"
            };
            officeReference.DigestMethod = _digestMethod;

            DataObject obj = new DataObject();
            obj.Id = "idOfficeObject";

            var props = new SignatureProperty("#idPackageSignature", "idOfficeV1Details");
            props.CreateSignatureInfo(SignerInformation);

            var propsXml = props.GetXMLDocument();
            obj.Data = propsXml.ChildNodes;

            signedXml.AddObject(obj);
            signedXml.AddReference(officeReference);

            return officeReference;
        }

        public Reference CreatePropertiesReference(ref ExcelSignedXml signedXml) 
        {
            Reference signedPropertiesReference = new()
            {
                Type = "http://uri.etsi.org/01903#SignedProperties",
                Uri = "#idSignedProperties"
            };
            XmlDsigC14NTransform c14Transform = new();

            signedPropertiesReference.AddTransform(c14Transform);
            signedPropertiesReference.DigestMethod = _digestMethod;

            DataObject signedProps = new DataObject();

            signedProps.LoadXml(_doc.DocumentElement);

            signedXml.AddObject(signedProps);
            signedXml.AddReference(signedPropertiesReference);

            return signedPropertiesReference;
        }

        public void SetDigestMethod(VbaSignatureHashAlgorithm algorithm)
        {
            switch (algorithm)
            {
                case VbaSignatureHashAlgorithm.MD5:
                    throw new InvalidOperationException("MD5 is not supported by excel or epplus for digital signatures. Please choose a different algorithm.");
                case VbaSignatureHashAlgorithm.SHA1:
                    _digestMethod = DigestMethods.SHA1;
                    break;
                case VbaSignatureHashAlgorithm.SHA256:
                    _digestMethod = DigestMethods.SHA256;
                    break;
                case VbaSignatureHashAlgorithm.SHA384:
                    _digestMethod = DigestMethods.SHA384;
                    break;
                case VbaSignatureHashAlgorithm.SHA512:
                    _digestMethod = DigestMethods.SHA512;
                    break;
            }
        }

        //internal void CreateSignature()
        //{
        //    byte[] certStore = CertUtil.GetSerializedCertStore(Certificate.RawData);
        //    if (Certificate == null)
        //    {
        //        SignaturePartUtil.DeleteParts(_part);
        //        return;
        //    }

        //    if (Certificate.HasPrivateKey == false)    //No signature. Remove any Signature part
        //    {
        //        var storeCert = CertUtil.GetCertificate(Certificate.Thumbprint);
        //        if (storeCert != null)
        //        {
        //            Certificate = storeCert;
        //        }
        //        else
        //        {
        //            SignaturePartUtil.DeleteParts(_part);
        //            return;
        //        }
        //    }

        //    using (var ms = RecyclableMemory.GetStream())
        //    {
        //        var bw = new BinaryWriter(ms);
        //        //Verifier = CertUtil.SignProject(project, this, Context);
        //        var cert = Verifier.Encode();
        //        var signatureBytes = CertUtil.CreateBinarySignature(ms, bw, certStore, cert);
        //        //_part = SignaturePartUtil.GetPart(project, this);
        //        _part.GetStream(FileMode.Create).Write(signatureBytes, 0, signatureBytes.Length);
        //    }
        //}
    }
}
