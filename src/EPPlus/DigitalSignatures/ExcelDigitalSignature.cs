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
using OfficeOpenXml.Drawing;
using System.IO;
using OfficeOpenXml.Drawing.EMF;

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

        ExcelSignatureLine _signatureLine;

        internal ExcelSignatureLine SignatureLine
        {
            get
            {
                return _signatureLine;
            }
            set
            {
                _signatureLine = value;
                SignatureLineSetupId = _signatureLine.SetupID;
            }
        }

        internal Guid? SignatureLineSetupId = null;

        internal string ValidSigLnImage;
        internal string InvalidSigLnImg;

        internal ExcelDigitalSignature(ExcelWorkbook wb, XmlNamespaceManager ns, ZipPackagePart part) : base(ns)
        {
            PartUri = part.Uri.OriginalString;
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
                if(string.IsNullOrEmpty(signatureProperty.sigInfo1.SetUpId) == false)
                {
                    SignatureLineSetupId = new Guid(signatureProperty.sigInfo1.SetUpId);

                    ValidSigLnImage = _doc.SelectSingleNode("//*[@Id='idValidSigLnImg']").InnerText;
                    InvalidSigLnImg = _doc.SelectSingleNode("//*[@Id='idInvalidSigLnImg']").InnerText;

                    //Could be made more effective if we only find the id string instead.
                    //Must load drawings to find SetupID in one of the shapes in one of the files.
                    //Worksheets must exist to load drawings.
                    var worksheets = _wb.Worksheets;
                    _wb.LoadAllDrawings("");
                }
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
            _originPart = wb._package.ZipPackage.GetPart(wb.SignatureOriginUri);
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
                CreateSignatureLineReferences(ref signedXml);

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

            if(SignatureLine != null)
            {
                props.sigInfo1.SetUpId = $"{{{SignatureLine.SetupID.ToString().ToUpper()}}}";

                if (SignatureLine.SignatureText != null)
                {
                    props.sigInfo1.SignatureText = SignatureLine.SignatureText;
                }

                if (SignatureLine.SignatureImage != null && SignatureLine.SignatureImage.Length > 0)
                {
                    var base64SLineString = Convert.ToBase64String(SignatureLine.SignatureImage);
                    props.sigInfo1.SignatureImage = base64SLineString;
                }
            }

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

        internal void CreateSignatureLineReferences(ref ExcelSignedXml signedXml)
        {
            if(SignatureLine != null)
            {
                Reference validImageReference = new()
                {
                    Type = "http://www.w3.org/2000/09/xmldsig#Object",
                    Uri = "#idValidSigLnImg"
                };
                validImageReference.DigestMethod = _digestMethod;
                Reference invalidImageReference = new()
                {
                    Type = "http://www.w3.org/2000/09/xmldsig#Object",
                    Uri = "#idInvalidSigLnImg"
                };
                invalidImageReference.DigestMethod = _digestMethod;

                DataObject validImageObject = new DataObject();
                DataObject invalidImageObject = new DataObject();


                var validTemplate = new SignatureLineTemplateValid();

                validTemplate.SuggestedSigner = SignatureLine.Signer;
                validTemplate.SuggestedTitle = SignatureLine.Title;
                validTemplate.SignedBy = Certificate.IssuerName.Name;
                validTemplate.SignText = SignatureLine.SignatureText;
                validTemplate.timeStamp.Text = DateTime.Now.ToString("yyyy-MM-dd");

                var invalidTemplate = new SignatureLineTemplateInvalid();
                invalidTemplate.SuggestedSigner = SignatureLine.Signer;
                invalidTemplate.SuggestedTitle = SignatureLine.Title;
                invalidTemplate.SignedBy = Certificate.IssuerName.Name;
                invalidTemplate.SignText = SignatureLine.SignatureText;

                validTemplate.Save(@"C:\epplusTest\Testoutput\ValidTemplateNew.emf");
                invalidTemplate.Save(@"C:\epplusTest\Testoutput\InvalidTemplateNew.emf");

                ValidSigLnImage = Convert.ToBase64String(validTemplate.GetBytes());
                InvalidSigLnImg = Convert.ToBase64String(invalidTemplate.GetBytes());

                XmlElement validElement = _doc.CreateElement("Object");
                validElement.SetAttribute("id", "idValidSigLnImg");
                validElement.InnerXml = ValidSigLnImage;

                XmlElement invalidElement = _doc.CreateElement("Object");
                invalidElement.SetAttribute("id", "idInvalidSigLnImg");
                invalidElement.InnerXml = InvalidSigLnImg;

                validImageObject.LoadXml(validElement);
                validImageObject.Id = "idValidSigLnImg";
                invalidImageObject.LoadXml(invalidElement);
                invalidImageObject.Id = "idInvalidSigLnImg";

                signedXml.AddObject(validImageObject);
                signedXml.AddReference(validImageReference);

                signedXml.AddObject(invalidImageObject);
                signedXml.AddReference(invalidImageReference);
            }
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

        public string HashAndEncodeBytes(byte[] temp)
        {
            using (var sha1Hash = SHA1.Create())
            {
                var hash = sha1Hash.ComputeHash(temp);
                return Convert.ToBase64String(hash);
            }
        }
    }
}
