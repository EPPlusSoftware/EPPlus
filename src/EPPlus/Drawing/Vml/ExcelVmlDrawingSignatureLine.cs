using System;
using System.Xml;

namespace OfficeOpenXml.Drawing.Vml
{
    /// <summary>
    /// Drawing object used for SignatureLines and SignatureLineStamps
    /// </summary>
    public class ExcelVmlDrawingSignatureLine : ExcelVmlDrawingBase
    {
        const string provIdStamp = "{000CD6A4-0000-0000-C000-000000000046}";
        const string provID = "{00000000-0000-0000-0000-000000000000}";
        internal Guid SetupID;

        internal ExcelVmlDrawingSignatureLine(XmlNode topNode, XmlNamespaceManager ns, Guid lineID) : base(topNode, ns)
        {
            SetupID = lineID;
            SetXmlNodeString("o:signatureline/@id", $"{{{SetupID.ToString().ToUpper()}}}");
            AlternativeText = "Microsoft Office Signature Line...";
            ShowSignDate = true;
            AllowComments = false;
            SigningInstructions = "Before signing this document, verify that the content you are signing is correct.";
        }

        internal ExcelVmlDrawingSignatureLine(XmlNode topNode, XmlNamespaceManager ns) : base(topNode, ns)
        {
            var idString = GetXmlNodeString("o:signatureline/@id");
            SetupID = new Guid(idString);
        }

        /// <summary>
        /// The suggested signer's name.
        /// </summary>
        public string Signer
        {
            get
            {
                var nodestring = GetXmlNodeString("o:signatureline/@o:suggestedsigner");
                return nodestring;
            }
            set
            {
                SetXmlNodeString("o:signatureline/@o:suggestedsigner", value);
            }
        }
        /// <summary>
        /// The suggested signers role or title e.g Developer.
        /// </summary>
        public string Title
        {
            get
            {
                var nodestring = GetXmlNodeString("o:signatureline/@o:suggestedsigner2");
                return nodestring;
            }
            set
            {
                SetXmlNodeString("o:signatureline/@o:suggestedsigner2", value);
            }
        }
        /// <summary>
        /// Suggested signers email.
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
        /// Instructions to the suggested signer.
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
                    line.SetAttribute("signinginstructionsset", "t");
                }
                else
                {
                    SetXmlNodeString("o:signatureline/@o:signinginstructions", value);
                }
            }
        }

        /// <summary>
        /// Allow signer to add comments such as commitmenttype and a "purpose" string
        /// </summary>
        public bool ShowSignDate
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
        /// Determines if signature allows comments such as commitment type and purpose
        /// </summary>
        public bool AllowComments
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
        internal bool IsStamp
        {
            get
            {
                return GetXmlNodeString("o:signatureline/@provid") == provIdStamp;
            }
            set
            {
                SetXmlNodeString("o:signatureline/@provid", value ? provIdStamp : provID);
                if(AlternativeText == "Stamp Signature Line..." || AlternativeText == "Microsoft Office Signature Line...")
                {
                    AlternativeText = value ? "Stamp Signature Line..." : "Microsoft Office Signature Line...";
                }
                Anchor = value ? "0, 0, 0, 0, 2, 0, 8, 0" : "0, 0, 0, 0, 4, 0, 6, 8";
            }
        }

        internal string ProvID
        {
            get
            {
                return GetXmlNodeString("o:signatureline/@provid");
            }
        }


        internal string RelId
        {
            get
            {
                return GetXmlNodeString("v:imagedata/@o:relid");
            }
            set
            {
                SetXmlNodeString("v:imagedata/@o:relid", value);
            }
        }

        ExcelVmlDrawingPosition _from = null;
        /// <summary>
        /// From position
        /// </summary>
        public ExcelVmlDrawingPosition From
        {
            get
            {
                if (_from == null)
                {
                    _from = new ExcelVmlDrawingPosition(NameSpaceManager, TopNode.SelectSingleNode("x:ClientData", NameSpaceManager), 0);
                }
                return _from;
            }
        }
        ExcelVmlDrawingPosition _to = null;
        /// <summary>
        /// To position
        /// </summary>
        public ExcelVmlDrawingPosition To
        {
            get
            {
                if (_to == null)
                {
                    _to = new ExcelVmlDrawingPosition(NameSpaceManager, TopNode.SelectSingleNode("x:ClientData", NameSpaceManager), 4);
                }
                return _to;
            }
        }

        const string vNameSpace = "urn:schemas-microsoft-com:vml";
        const string oNameSpace = "urn:schemas-microsoft-com:office:office";
        const string xNameSpace = "urn:schemas-microsoft-com:office:excel";

        internal static void CreateFormulaElementAsChildOf(XmlNode node)
        {
            var doc = node.OwnerDocument;

            var formulaElement = doc.CreateElement("v", "formulas", vNameSpace);
            node.AppendChild(formulaElement);

            CreateAndSetFormulaElementOnNode(formulaElement, doc, "if lineDrawn pixelLineWidth 0");
            CreateAndSetFormulaElementOnNode(formulaElement, doc, "sum @0 1 0");
            CreateAndSetFormulaElementOnNode(formulaElement, doc, "sum 0 0 @1");
            CreateAndSetFormulaElementOnNode(formulaElement, doc, "prod @2 1 2");
            CreateAndSetFormulaElementOnNode(formulaElement, doc, "prod @3 21600 pixelWidth");
            CreateAndSetFormulaElementOnNode(formulaElement, doc, "prod @3 21600 pixelHeight");
            CreateAndSetFormulaElementOnNode(formulaElement, doc, "sum @0 0 1");
            CreateAndSetFormulaElementOnNode(formulaElement, doc, "prod @6 1 2");
            CreateAndSetFormulaElementOnNode(formulaElement, doc, "prod @7 21600 pixelWidth");
            CreateAndSetFormulaElementOnNode(formulaElement, doc, "sum @8 21600 0");
            CreateAndSetFormulaElementOnNode(formulaElement, doc, "prod @7 21600 pixelHeight");
            CreateAndSetFormulaElementOnNode(formulaElement, doc, "sum @10 21600 0");
        }

        static void CreateAndSetFormulaElementOnNode(XmlElement formulaParentNode, XmlDocument document, string formula)
        {
            var f1 = document.CreateElement("v", "f", vNameSpace);
            f1.SetAttribute("eqn", formula);
            formulaParentNode.AppendChild(f1);
        }
    }
}
