using OfficeOpenXml.DigitalSignatures;
using OfficeOpenXml.Drawing.EMF;
using OfficeOpenXml.Packaging;
using System.Xml;
using OfficeOpenXml.Drawing.Vml;
using System;
using System.IO;

namespace OfficeOpenXml.Drawing
{
    public class ExcelSignatureLineStamp : ExcelVmlDrawingSignatureLine
    {
        protected ExcelImage _signatureImage = null;
        protected string _signatureText = "";

        /// <summary>
        /// The Signature itself.
        /// Note that setting SignatureImage will erase SignatureText and vice-versa
        /// </summary>
        public virtual ExcelImage SignatureImage
        {
            get
            {
                return _signatureImage;
            }
            set
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

        ExcelDigitalSignature Signature = null;
        SignatureLineTemplateEmfBase Emf;

        internal string InvalidSigLnImg;
        internal string ValidSigLnImage;

        internal ExcelSignatureLineStamp(ExcelWorksheet ws, XmlNode topNode, XmlNamespaceManager ns, Guid lineId) : base(topNode, ns, lineId)
        {
            IsStamp = true;
        }

        internal ExcelSignatureLineStamp(ExcelWorksheet ws, XmlNode topNode, XmlNamespaceManager ns) : base(topNode, ns)
        {
            IsStamp = true;
        }

        internal void SaveMedia(ZipPackagePart part)
        {
            Emf = IsStamp ? new SignatureLineTemplateEmfStamp() : new SignatureLineTemplateEmf();
            Emf.SuggestedSigner = Signer;
            Emf.SuggestedTitle = Title;

            Emf.Save("C:\\epplusTest\\Testoutput\\image1Generated.emf");

            //Note: Intentionally not disposed.
            MemoryStream ms = (MemoryStream)part.GetStream(FileMode.Create, FileAccess.Write);
            Emf.SaveToStream(ms);
        }

        internal void SaveSignatureLineWithDigitalSignature(string signerName)
        {
            Emf.SaveTemplateProperties([signerName, _signatureText]);

            if(SignatureImage != null && SignatureImage.ImageBytes.Length > 0)
            {
                Emf.SaveImage(SignatureImage.ImageBytes);
            }

            if (ShowSignDate)
            {
                Emf.TimeStamp = DateTime.Now.ToString("yyyy-MM-dd");
            }

            ValidSigLnImage = Convert.ToBase64String(Emf.GetBytes());
            Emf.Save("C:\\epplusTest\\Testoutput\\ValidTemplateNew.emf");

            Emf.timeStamp.Text = "";
            Emf.InsertInvalidRecords();

            InvalidSigLnImg = Convert.ToBase64String(Emf.GetBytes());
            Emf.Save("C:\\epplusTest\\Testoutput\\InvalidTemplateNew.emf");
        }
    }
}
