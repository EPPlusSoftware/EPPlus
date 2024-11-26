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

namespace OfficeOpenXml.DigitalSignatures
{
    public class ExcelDigitalSignature : XmlHelper
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

        bool wasRead = false;
        private SignatureProperty signatureProperty;

        /// <summary>
        /// Details about the signer of a DigitalSignature, such as role, title, address etc.
        /// </summary>
        public AdditionalSignatureInfo SigningInformation = new AdditionalSignatureInfo();

        /// <summary>
        /// Reason for signing the document-
        /// </summary>
        public string PurposeForSigning = "";
        /// <summary>
        /// Commitment Type.
        /// </summary>
        public CommitmentType CommitmentTyping = CommitmentType.None;

        internal Guid SetupId;
        //private bool _verified = false;

        /// <summary>
        /// Wheter the Signature is verified to be valid
        /// </summary>
        public bool Verified 
        {
            get
            {
                SignedXml signedXml = new SignedXml(_doc);
                return signedXml.CheckSignatureReturningKey(out AsymmetricAlgorithm pubKey);
            }
        }

        QualifyingProperties qualifyingProperties;
        ExcelSignatureLineStamp _signatureLine;

        internal ExcelSignatureLineStamp SignatureLine
        {
            get
            {
                return _signatureLine;
            }
            set
            {
                _signatureLine = value;
            }
        }

        internal string ValidSigLnImage;
        internal string InvalidSigLnImg;
        internal string SignatureText;
        internal string SignatureImage;

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
                signatureProperty = new SignatureProperty((XmlElement)officeObj, SigningInformation);
                PurposeForSigning = signatureProperty.sigInfo1.SignatureComments;

                if (string.IsNullOrEmpty(signatureProperty.sigInfo1.SetUpId) == false)
                {
                    ValidSigLnImage = _doc.SelectSingleNode("//*[@Id='idValidSigLnImg']").InnerText;
                    InvalidSigLnImg = _doc.SelectSingleNode("//*[@Id='idInvalidSigLnImg']").InnerText;

                    SignatureImage = signatureProperty.sigInfo1.SignatureImage;
                    SignatureText = signatureProperty.sigInfo1.SignatureText;

                    //Could be made more effective if we only find the  id string via part instead.
                    //Must load drawings to find SetupID in one of the shapes in one of the files.
                    //Worksheets must exist to load drawings.
                    var worksheets = _wb.Worksheets;
                    _wb.LoadAllVmlDrawings("");

                    SetupId = new Guid(signatureProperty.sigInfo1.SetUpId);
                    _signatureLine = _wb.GetSignatureLineStamp(SetupId);

                    _signatureLine.ValidSigLnImage = ValidSigLnImage;
                    _signatureLine.InvalidSigLnImg = InvalidSigLnImg;

                    if (string.IsNullOrEmpty(SignatureText) == false)
                    {
                        _signatureLine.AsSignatureLine.SignatureText = SignatureText;
                    }

                    if(string.IsNullOrEmpty(SignatureImage) == false)
                    {
                        var emfBytes = Convert.FromBase64String(SignatureImage);
                        SignatureLine.ReadEmfExtractImage(emfBytes);
                    }
                }
            }

            var typeQualifiers = new List<string>();
            var signedPropertiesNode = _doc.SelectSingleNode("//*[@Id='idSignedProperties']");
            qualifyingProperties = new QualifyingProperties((XmlElement)signedPropertiesNode, SigningInformation, typeQualifiers, ref CommitmentTyping);

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

        internal void Save()
        {
            if (Verified == false && Certificate != null)
            {
                if (SignatureLine != null)
                {
                    SignatureLine.SaveSignatureLineWithDigitalSignature(Certificate.IssuerName.Name.Substring(3));
                    ValidSigLnImage = SignatureLine.ValidSigLnImage;
                    InvalidSigLnImg = SignatureLine.InvalidSigLnImg;
                }

                var signatureComments = new List<string>
                {
                    PurposeForSigning
                };

                qualifyingProperties = new QualifyingProperties
                    ("xd", Certificate, CommitmentTyping, signatureComments, SigningInformation);

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

                var declaration = outPutDoc.CreateXmlDeclaration("1.0", "UTF-8", "");
                outPutDoc.InsertBefore(declaration, node);

                var stream = _part.GetStream();
                stream.Position = 0;

                _doc = outPutDoc;

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

        internal Reference CreatePackageReference(ref ExcelSignedXml signedXml)
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

        internal Reference CreateOfficeReference(ref ExcelSignedXml signedXml)
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
            props.CreateSignatureInfo(SigningInformation, PurposeForSigning);

            if(SignatureLine != null)
            {
                props.sigInfo1.SetUpId = $"{{{SignatureLine.SetupID.ToString().ToUpper()}}}";
                props.sigInfo1.SignatureProviderID = SignatureLine.ProvID;

                if(SignatureLine is ExcelSignatureLine)
                {
                    var sigLine = SignatureLine as ExcelSignatureLine;
                    if (sigLine.SignatureText != null)
                    {
                        props.sigInfo1.SignatureText = sigLine.SignatureText;
                    }
                }

                if (SignatureLine.SignatureImage != null && SignatureLine.SignatureImage.ImageBytes.Length > 0)
                {
                    var base64SLineString = Convert.ToBase64String(SignatureLine.SignatureImage.ImageBytes);
                    props.sigInfo1.SignatureImage = base64SLineString;
                }
            }

            var propsXml = props.GetXMLDocument();
            obj.Data = propsXml.ChildNodes;

            signedXml.AddObject(obj);
            signedXml.AddReference(officeReference);

            return officeReference;
        }

        internal Reference CreatePropertiesReference(ref ExcelSignedXml signedXml) 
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

                XmlElement validElement = _doc.CreateElement("Object");
                validElement.SetAttribute("id", "idValidSigLnImg");
                validElement.InnerXml = SignatureLine.ValidSigLnImage;

                XmlElement invalidElement = _doc.CreateElement("Object");
                invalidElement.SetAttribute("id", "idInvalidSigLnImg");
                invalidElement.InnerXml = SignatureLine.InvalidSigLnImg;

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
