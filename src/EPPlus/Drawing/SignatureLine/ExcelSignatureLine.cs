using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Xml;
using OfficeOpenXml.DigitalSignatures;
using OfficeOpenXml.Drawing.EMF;
using OfficeOpenXml.Packaging;
using OfficeOpenXml.Drawing.Vml;
using System.IO;

namespace OfficeOpenXml.Drawing
{
    /// <summary>
    /// Signature line which can contain either text or an image as the signature.
    /// </summary>
    public class ExcelSignatureLine : ExcelSignatureLineStamp
    {
        /// <summary>
        /// The Signature itself.
        /// Cannot be set if IsStamp is true.
        /// Note that setting SignatureText will erase SignatureImage and vice-versa.
        /// </summary>
        public string SignatureText
        {
            get
            {
                return _signatureText;
            }
            private set
            {
                if (string.IsNullOrEmpty(value))
                {
                    throw new InvalidOperationException($"Cannot set SignatureText of SignatureLine object {this} to null or empty.");
                }
                _signatureText = value;
                _signatureImage = null;
            }
        }

        internal override void CheckSignature()
        {
            if ((SignatureImage == null || SignatureImage.ImageBytes.Length <= 0) && string.IsNullOrEmpty(SignatureText))
            {
                throw new InvalidOperationException($"SignatureLine {this} is invalid. Cannot sign without a Signature. Please add a SignatureImage or SignatureText first.");
            }
        }

        internal ExcelSignatureLine(ExcelWorksheet ws, XmlNode topNode, XmlNamespaceManager ns, Guid lineId) : base(ws, topNode, ns, lineId)
        {
            _signatureLineType = eSignatureLineType.SignatureLine;
            IsStamp = false;
        }

        internal ExcelSignatureLine(ExcelWorksheet ws, XmlNode topNode, XmlNamespaceManager ns) : base(ws, topNode, ns)
        {
            _signatureLineType = eSignatureLineType.SignatureLine;
        }


        /// <summary>
        /// Sign the signatureline with a new digital signature.
        /// </summary>
        /// <param name="signatureText">Cannot be null or empty</param>
        /// <param name="certificate"></param>
        /// <param name="cType"></param>
        /// <param name="purposeForSigning"></param>
        /// <returns></returns>
        public ExcelDigitalSignature Sign(X509Certificate2 certificate, string signatureText, CommitmentType cType = CommitmentType.None, string purposeForSigning = "")
        {
            SignatureText = signatureText;
            CheckSignature();
            return Sign(certificate, cType, purposeForSigning);
        }
    }
}
