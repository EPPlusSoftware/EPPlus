using OfficeOpenXml.DigitalSignatures;
using OfficeOpenXml.Drawing.EMF;
using OfficeOpenXml.Packaging;
using System.Xml;
using OfficeOpenXml.Drawing.Vml;
using System;
using System.IO;

namespace OfficeOpenXml.Drawing
{
    public class ExcelSignatureLine : ExcelVmlDrawingSignatureLine
    {
        /// <summary>
        /// The signature itself as text.
        /// Has no point in a signature stamp
        /// </summary>
        public string SignatureText = "";

        /// <summary>
        /// Note that while SignatureText and SignatureImage are Allowed to both exist.
        /// SignatureImage will override SignatureText visually if it exists.
        /// </summary>
        public byte[] SignatureImage;

        ExcelDigitalSignature Signature = null;
        SignatureLineTemplateEmf Emf;

        internal string InvalidSigLnImg;
        internal string ValidSigLnImage;

        internal ExcelSignatureLine(ExcelWorksheet ws, XmlNode topNode, XmlNamespaceManager ns, Guid lineId) : base(topNode, ns, lineId)
        {
        }

        internal ExcelSignatureLine(ExcelWorksheet ws, XmlNode topNode, XmlNamespaceManager ns) : base(topNode, ns)
        {
        }

        internal void SaveMedia(ZipPackagePart part)
        {
            Emf = new SignatureLineTemplateEmf(IsStamp);
            Emf.SuggestedSigner = Signer;
            Emf.SuggestedTitle = Title;
            Emf.timeStamp.Text = "";
            if(Emf.signTextObject != null)
            {
                Emf.SignText = "";
            }

            //Note: Intentionally not disposed.
            MemoryStream ms = (MemoryStream)part.GetStream(FileMode.Create, FileAccess.Write);
            Emf.SaveToStream(ms);

            Emf.Save("C:\\epplusTest\\Testoutput\\image1Generated.emf");
        }

        internal void SaveSignatureLineWithDigitalSignature(string signerName)
        {
            if (IsStamp == false)
            {
                Emf.SignText = SignatureText;
                Emf.SignedBy = signerName;
            }
            else
            {
                //Stamps replace suggested with signername
                Emf.SuggestedSigner = signerName;
            }

            if (ShowSignDate)
            {
                Emf.timeStamp.Text = DateTime.Now.ToString("yyyy-MM-dd");
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
