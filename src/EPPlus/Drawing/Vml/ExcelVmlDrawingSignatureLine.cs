using OfficeOpenXml.Drawing.EMF;
using OfficeOpenXml.Drawing.Interfaces;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Statistical;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Text;
using OfficeOpenXml.Packaging;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection.Emit;
using System.Text;
using System.Xml;

namespace OfficeOpenXml.Drawing.Vml
{
    public class ExcelVmlDrawingSignatureLine : ExcelVmlDrawingBase
    {
        internal SignatureLineEmf Emf;
        const string provIdStamp = "{000CD6A4-0000-0000-C000-000000000046}";
        const string provID = "{00000000-0000-0000-0000-000000000000}";
        internal Guid LineId;

        internal ExcelVmlDrawingSignatureLine(XmlNode topNode, XmlNamespaceManager ns, Guid lineID) : base(topNode, ns)
        {
            Emf = new SignatureLineEmf();
            Emf.SignerName = Signer;
            Emf.SignerTitle = Title;
            LineId = lineID;
            SetXmlNodeString("o:signatureline/@id", LineId.ToString());
            AlternativeText = "Microsoft Office Signature Line...";
        }

        /// <summary>
        /// The suggested name for the signer
        /// </summary>
        public string Signer
        {
            get
            {
                var nodestring = GetXmlNodeString("o:signatureline/@o:suggestedsigner");
                Emf.SignerName = nodestring;
                return nodestring;
            }
            set
            {
                SetXmlNodeString("o:signatureline/@o:suggestedsigner", value);
                Emf.SignerName = value;
            }
        }
        /// <summary>
        /// The suggested signers role or title e.g Developer
        /// </summary>
        public string Title
        {
            get
            {
                var nodestring = GetXmlNodeString("o:signatureline/@o:suggestedsigner2");
                Emf.SignerName = nodestring;
                return nodestring;
            }
            set
            {
                SetXmlNodeString("o:signatureline/@o:suggestedsigner2", value);
                Emf.SignerTitle = value;
            }
        }
        /// <summary>
        /// Suggested signers email
        /// </summary>
        public string Email
        {
            get
            {
                return GetXmlNodeString("o:signatureline/@o:suggestedsigneremail");
            }
            set
            {
                SetXmlNodeString("o:signatureline/@o:suggestedsigneremail", value);
            }
        }
        /// <summary>
        /// Instructions for the suggested signer
        /// </summary>
        public string SigningInstructions
        {
            get
            {
                return GetXmlNodeString("o:signatureline/@o:signinginstructions");
            }
            set
            {
                if(string.IsNullOrEmpty(GetXmlNodeString("o:signatureline/@o:signinginstructions")))
                {
                    SetXmlNodeString("o:signatureline/@o:signinginstructions", value);
                    var line = (XmlElement)TopNode.SelectSingleNode("o:signatureline", NameSpaceManager);
                    line.SetAttribute("allowcomments", "t");
                    line.SetAttribute("signinginstructionsset", "t");
                }
                else
                {
                    SetXmlNodeString("o:signatureline/@o:signinginstructions", value);
                }
            }
        }

        /// <summary>
        /// Instructions for the suggested signer
        /// </summary>
        public bool AllowComments
        {
            get
            {
                return GetXmlNodeBool("o:signatureline/@showsigndate");
            }
            set
            {
                if (string.IsNullOrEmpty(GetXmlNodeString("o:signatureline/@showsigndate")))
                {
                    SetXmlNodeBoolVml("o:signatureline/@showsigndate", value);
                }
                else
                {
                    SetXmlNodeBoolVml("o:signatureline/@showsigndate", value);
                }
            }
        }

        /// <summary>
        /// Instructions for the suggested signer
        /// </summary>
        public bool ShowSignDate
        {
            get
            {
                return GetXmlNodeBool("o:signatureline/@allowcomments");
            }
            set
            {
                if (string.IsNullOrEmpty(GetXmlNodeString("o:signatureline/@allowcomments")))
                {
                    SetXmlNodeBoolVml("o:signatureline/@allowcomments", value);
                }
                else
                {
                    SetXmlNodeBoolVml("o:signatureline/@allowcomments", value);
                }
            }
        }

        /// <summary>
        /// True if digital signature is stamp type. False by default
        /// </summary>
        public bool IsStamp
        {
            get
            {
                return GetXmlNodeString("o:signatureline/@provId") == provIdStamp;
            }
            set
            {
                SetXmlNodeString("o:signatureline/@provId", value ? provIdStamp : provID);
                AlternativeText = value ? "Stamp Signature Line..." : "Microsoft Office Signature Line...";
                Anchor = value ? "0, 0, 0, 0, 2, 0, 8, 0" : "0, 0, 0, 0, 4, 0, 6, 8";
            }
        }
    }
}
