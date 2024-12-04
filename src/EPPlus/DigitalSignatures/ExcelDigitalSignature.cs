using System;
using System.Security.Cryptography.X509Certificates;
using OfficeOpenXml.Constants;
using OfficeOpenXml.Packaging;
using System.Security.Cryptography;
using System.Security.Cryptography.Xml;
using System.Xml;
using OfficeOpenXml.DigitalSignatures.XAdES;
using System.Collections.Generic;
using System.Linq;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Utils;
using System.Runtime.ConstrainedExecution;
using OfficeOpenXml.FormulaParsing.Excel.Functions.MathFunctions;

namespace OfficeOpenXml.DigitalSignatures
{
    /// <summary>
    /// Digital Signature class to sign a workbook
    /// </summary>
    public class ExcelDigitalSignature : XmlHelper
    {
        internal ZipPackagePart _part;
        ZipPackagePart _originPart;
        ExcelWorkbook _wb;
        DigSigManifest _manifest;

        private X509Certificate2 _cert = null;

        /// <summary>
        /// The Certificate used to sign this digital signature
        /// </summary>
        public X509Certificate2 Certificate
        {
            get
            {
                return _cert;
            }
            set
            {
                _cert = value;
            }
        }

        const string _originPartUri = @"/_xmlsignatures/origin.sigs";
        const string relType = "http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/signature";
        const string relTypeOrigin = "http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/origin";
        private const string PartUriBase = @"/_xmlsignatures/sig{0}.xml";
        internal string PartUri = "";

        internal string _digestMethod { get; private set; } = DigestMethods.SHA1;
        internal string _signatureMethod { get; private set; } = SignedXml.XmlDsigRSASHA1Url;
        string _referenceType = "http://www.w3.org/2000/09/xmldsig#Object";

        XmlDocument _doc;

        bool wasRead = false;
        private SignatureProperty signatureProperty;

        /// <summary>
        /// Details about the signer of a DigitalSignature, such as role, title, address etc.
        /// </summary>
        public AdditionalSignatureInfo Details = new AdditionalSignatureInfo();

        /// <summary>
        /// Reason for signing the document-
        /// </summary>
        public string PurposeForSigning = "";
        /// <summary>
        /// Commitment Type.
        /// </summary>
        public CommitmentType CommitmentTyping = CommitmentType.None;

        private Guid SetupId;
        //private bool _verified = false;

        DigSigManifest readManifest = null;

        /// <summary>
        /// Whether the Signature was valid when the file was read/saved
        /// Is dirty and outdated until after package is saved if any changes have been made to package files.
        /// </summary>
        public bool IsValid
        {
            get
            {
                if(_doc != null)
                {
                    SignedXml signedXml = new SignedXml(_doc);

                    var node = _doc.GetElementsByTagName("Signature")[0];
                    if(node != null)
                    {
                        signedXml.LoadXml((XmlElement)node);

                        return signedXml.CheckSignatureReturningKey(out AsymmetricAlgorithm pubKey);
                    }
                }
                return false;
            }
        }

        QualifyingProperties _qualifyingProperties;
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
        internal string SigLnImage;

        //Read
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
            var isVerified = IsValid;

            //Read signing method
            SignedXml signedXml = new();
            signedXml.LoadXml(_doc.DocumentElement);
            _signatureMethod = signedXml.SignedInfo.SignatureMethod;

            //Read digest method (Assume inital digest is same for all other)
            Reference firstRef = (Reference)signedXml.SignedInfo.References.ToArray()[0];
            _digestMethod = firstRef.DigestMethod;

            var readAlgorithm = DigestMethods.GetHashAlgorithmByDigest(_digestMethod);
            if(readAlgorithm != null)
            {
                _wb._package.ZipPackage.hashAlgorithm = readAlgorithm.Value;
            }
            else
            {
                //Unknown algorithm. Throw? Apply default? Will throw on save
            }

            var packageObj = _doc.GetElementsByTagName("Manifest");
            readManifest = new DigSigManifest(packageObj[0]);

            var officeObj = _doc.SelectSingleNode("//*[@Id='idOfficeV1Details']");

            if(officeObj != null)
            {
                signatureProperty = new SignatureProperty((XmlElement)officeObj, Details);
                PurposeForSigning = signatureProperty.sigInfo1.SignatureComments;

                if (string.IsNullOrEmpty(signatureProperty.sigInfo1.SetUpId) == false)
                {
                    ValidSigLnImage = _doc.SelectSingleNode("//*[@Id='idValidSigLnImg']").InnerText;
                    InvalidSigLnImg = _doc.SelectSingleNode("//*[@Id='idInvalidSigLnImg']").InnerText;

                    SigLnImage = signatureProperty.sigInfo1.SignatureImage;
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

                    if(string.IsNullOrEmpty(SigLnImage) == false)
                    {
                        var emfBytes = Convert.FromBase64String(SigLnImage);
                        SignatureLine.ReadEmfExtractImage(emfBytes);
                    }
                }
            }

            var typeQualifiers = new List<string>();
            var signedPropertiesNode = _doc.SelectSingleNode("//*[@Id='idSignedProperties']");
            _qualifyingProperties = new QualifyingProperties((XmlElement)signedPropertiesNode, Details, typeQualifiers, ref CommitmentTyping);

            string keyInfo = _doc.GetElementsByTagName("KeyInfo")[0].InnerText;
            string serialInFile = _qualifyingProperties.SignedProps.SignatureProps.Serial;


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
            SetDigestMethod(DigitalSignatureHashAlgorithm.SHA1);
        }

        internal void Save()
        {
            if (Certificate != null)
            {
#if NET35
#else
                var someKey = Certificate.GetRSAPrivateKey();
#endif

                //if there is no read manifest then the manifest has changed
                bool manifestChanged = true;
  
                _manifest = new DigSigManifest(_wb._package.ZipPackage.XmlManifest, _wb._package.ZipPackage.hashAlgorithm);

                if (readManifest != null)
                {
                    var newManifest = _manifest.GetDoc().OuterXml;
                    var oldManifest = readManifest.GetDoc().OuterXml;

                    manifestChanged = !newManifest.Equals(oldManifest);
                }

                if (IsValid == false | manifestChanged)
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

                    _qualifyingProperties = new QualifyingProperties
                        ("xd", Certificate, CommitmentTyping, signatureComments, Details);

                    //Set digestmethod and hash for certDigest
                    _qualifyingProperties.SignedProps.SignatureProps.Algorithm = _digestMethod;
                    _qualifyingProperties.SignedProps.SignatureProps.HashCert(_wb._package.ZipPackage.hashAlgorithm);


                    var docTest = _qualifyingProperties.GetDocument();
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
                    signedXml.SignedInfo.SignatureMethod = _signatureMethod;

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

            var packageObjProps = new SignatureProperty("#idPackageSignature", "idSignatureTime", DateTime.Now);

            var packageObjectDoc = new XmlDocument();

            XmlDocument docManifest = _manifest.GetDoc();

            //XmlDocument docManifest = _wb._package.ZipPackage.Manifest.GetDoc();
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
            props.CreateSignatureInfo(Details, PurposeForSigning);

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
                    props.sigInfo1.SignatureImage = SignatureLine.SigLnImage;
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

        /// <summary>
        /// Set the digest method/Hash Algorithm for the package
        /// Note: All Digital Signatures in the package will use the latest algorithm.
        /// </summary>
        /// <param name="algorithm"></param>
        /// <exception cref="InvalidOperationException"></exception>
        public void SetDigestMethod(DigitalSignatureHashAlgorithm algorithm)
        {
            _wb._package.ZipPackage.hashAlgorithm = algorithm;
            _digestMethod = DigestMethods.GetDigestMethod(algorithm);
            _signatureMethod = DigestMethods.GetSignatureMethod(algorithm);
        }

        internal string GetOuterXml()
        {
            if(_doc != null)
            {
                return _doc.OuterXml;
            }
            else
            {
                return "";
            }
        }
    }
}
