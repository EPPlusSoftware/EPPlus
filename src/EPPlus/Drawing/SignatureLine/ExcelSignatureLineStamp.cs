using OfficeOpenXml.DigitalSignatures;
using OfficeOpenXml.Drawing.EMF;
using OfficeOpenXml.Packaging;
using System.Xml;
using OfficeOpenXml.Drawing.Vml;
using System;
using System.IO;
using System.Security.Cryptography.X509Certificates;
using System.Linq;

namespace OfficeOpenXml.Drawing
{
    /// <summary>
    /// Signature line stamp which contains a SignatureImage as the signature.
    /// </summary>
    public class ExcelSignatureLineStamp : ExcelVmlDrawingSignatureLine
    {
        private protected ExcelImage _signatureImage = null;
        private protected string _signatureText = "";
        internal ExcelWorkbook wb;
        internal ExcelWorksheet _ws;

        private protected eSignatureLineType _signatureLineType = eSignatureLineType.Stamp;
        ePictureType[] restrictedTypes = Enum.GetValues(typeof(ePictureType)).Cast<ePictureType>().Where(x => x != ePictureType.Bmp).ToArray();
        /// <summary>
        /// The type of signatureline
        /// </summary>
        public eSignatureLineType SignatureLineType
        {
            get
            {
                return _signatureLineType;
            }
        }

        /// <summary>
        /// The Signature itself.
        /// Note that setting SignatureImage will erase SignatureText and vice-versa
        /// Must be .bmp format.
        /// </summary>
        public virtual ExcelImage SignatureImage
        {
            get
            {
                return _signatureImage;
            }
            internal set
            {
                if (value.Type != ePictureType.Bmp)
                {
                    throw new InvalidOperationException($"SignatureImage must be of type {ePictureType.Bmp}. SignatureLine object {this} cannot be set.");
                }
                //Technically we Could allow this and throw only on save but to allow it here is to allow a more confusing error later.
                if (value == null && value.ImageBytes.Length <= 0)
                {
                    throw new InvalidOperationException($"SignatureImage.ImageBytes must be > 0 and not Null. Imagebytes.Length where: {SignatureImage.ImageBytes.Length}.\n" +
                        $"SignatureLine{this} SignatureImage cannot be set.");
                }
                _signatureImage = value;
                _signatureText = "";
            }
        }

        /// <summary>
        /// The digital signature this signatureline is linked to one.
        /// Must be set via Sign methods.
        /// </summary>
        public ExcelDigitalSignature DigitalSignature
        {
            get
            {
                return wb.DigitalSignatures.GetSignatureBySignatureLineGuid(SetupID);
            }
        }
        SignatureLineTemplateEmfBase Emf;

        internal string InvalidSigLnImg;
        internal string ValidSigLnImage;
        internal string SigLnImage;

        internal ExcelSignatureLineStamp(ExcelWorksheet ws, XmlNode topNode, XmlNamespaceManager ns, Guid lineId) : base(topNode, ns, lineId, ws)
        {
            _ws = ws;
            wb = ws.Workbook;
            IsStamp = true;

            //Setting default size
            From.Column = 0;
            To.Column = 2;
            From.Row = 0;
            To.Row = 8;
        }

        internal ExcelSignatureLineStamp(ExcelWorksheet ws, XmlNode topNode, XmlNamespaceManager ns) : base(topNode, ns, ws)
        {
            _ws = ws;
            wb = ws.Workbook;
        }

        internal void SaveMedia(ZipPackagePart part)
        {
            Emf = _signatureLineType == eSignatureLineType.Stamp ? new SignatureLineTemplateEmfStamp() : new SignatureLineTemplateEmf();
            Emf.SuggestedSigner = Signer;
            Emf.SuggestedTitle = Title;

            if (To.Row <= From.Row)
            {
                throw new InvalidDataException($"To.Row property '{To.Row}' cannot be less than or equal to From.Row property '{From.Row}'. Signature line with id '{SetupID}' will not be visible.");
            }
            else if(To.Column <= From.Column)
            {
                throw new InvalidDataException($"To.Column property '{To.Column}' cannot be less than or equal to From.Column property '{From.Column}'. signature line '{SetupID}' will not be visible.");
            }

            //Note: Intentionally not disposed.
            MemoryStream ms = (MemoryStream)part.GetStream(FileMode.Create, FileAccess.Write);
            Emf.SaveToStream(ms);
        }

        internal void SaveSignatureLineWithDigitalSignature(string signerName)
        {
            Emf.SaveTemplateProperties([signerName, _signatureText]);

            if(SignatureImage != null && SignatureImage.ImageBytes.Length > 0)
            {
                if(SignatureImage.Type == ePictureType.Bmp)
                {
                    Emf.SaveImage(SignatureImage.ImageBytes);
                    SigLnImage = Convert.ToBase64String(Emf.EmfSignatureImage.GetBytes());
                }
                else
                {
                    throw new InvalidOperationException($"SignatureImage must be .bmp format/type. SignatureImage was of type {SignatureImage.Type}");
                }
            }

            if (ShowSignDate)
            {
                Emf.TimeStamp = DateTime.Now.ToString("yyyy-MM-dd");
            }

            ValidSigLnImage = Convert.ToBase64String(Emf.GetBytes());

            Emf.timeStamp.Text = "";
            Emf.InsertInvalidRecords();

            InvalidSigLnImg = Convert.ToBase64String(Emf.GetBytes());
        }

        internal virtual void CheckSignature()
        {
            if(SignatureImage == null || SignatureImage.ImageBytes.Length <= 0)
            {
                throw new InvalidOperationException($"SignatureLine {this} is invalid. Cannot sign without a Signature. Please add a SignatureImage first.");
            }
            SignatureImage.SetRestrictedTypes(restrictedTypes);
        }

        internal virtual void ReadEmfExtractImage(byte[] emfBytes)
        {
            Emf = _signatureLineType == eSignatureLineType.Stamp ? new SignatureLineTemplateEmfStamp(emfBytes) : new SignatureLineTemplateEmf(emfBytes);
            SignatureImage = new ExcelImage(Emf.GetBitmapBytes(), ePictureType.Bmp);
            SigLnImage = Convert.ToBase64String(Emf.EmfSignatureImage.GetBytes());
        }

        /// <summary>
        /// Sign the signatureline with a new digital signature.
        /// The image must be .bmp format
        /// </summary>
        /// <param name="image">Must be in .bmp format</param>
        /// <param name="certificate">Certificate for creating the digital signature</param>
        /// <returns></returns>
        public ExcelDigitalSignature SignWithImage(X509Certificate2 certificate, ExcelImage image)
        {
            SignatureImage = image;
            CheckSignature();
            return Sign(certificate);
        }

        /// <summary>
        /// </summary>
        /// <param name="certificate"></param>
        /// <returns></returns>
        private protected ExcelDigitalSignature Sign(X509Certificate2 certificate)
        {
            var digSig = wb.DigitalSignatures.Add(certificate);
            digSig.SignatureLine = this;
            return digSig;
        }

        /// <summary>
        /// Return this as signatureline
        /// </summary>
        public ExcelSignatureLine AsSignatureLine { get { return this as ExcelSignatureLine; } }
    }
}
