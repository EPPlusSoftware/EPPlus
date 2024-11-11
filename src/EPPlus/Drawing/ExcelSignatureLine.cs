using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;

namespace OfficeOpenXml.Drawing
{
    public class ExcelSignatureLine : ExcelSignatureLineStamp
    {
        private string _signatureText = "";

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
            set
            {
                if (string.IsNullOrEmpty(value))
                {
                    throw new InvalidOperationException($"Cannot set SignatureText of SignatureLine object {this} to null or empty.");
                }
                if (IsStamp)
                {
                    throw new InvalidOperationException($"Cannot set SignatureText on a SignatureStamp object.");
                }
                _signatureText = value;
                _signatureImage = null;
            }
        }

        internal ExcelSignatureLine(ExcelWorksheet ws, XmlNode topNode, XmlNamespaceManager ns, Guid lineId) : base(ws, topNode, ns, lineId)
        {
            IsStamp = false;
        }

        internal ExcelSignatureLine(ExcelWorksheet ws, XmlNode topNode, XmlNamespaceManager ns) : base(ws, topNode, ns)
        {
            IsStamp = false;
        }
    }
}
