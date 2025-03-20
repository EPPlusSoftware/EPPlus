using OfficeOpenXml.FormulaParsing.Excel.Functions.MathFunctions;
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
        ExcelWorksheet _ws;

        /// <summary>
        /// Id of signature line
        /// </summary>
        public Guid SetupID { get; internal set; }

        const string StyleAttrPath = "@style";

        string styleString;
        //string styleWidth;
        //string styleHeight;

        internal ExcelVmlDrawingSignatureLine(XmlNode topNode, XmlNamespaceManager ns, Guid lineID, ExcelWorksheet ws) : base(topNode, ns)
        {
            SetupID = lineID;
            SetXmlNodeString("o:signatureline/@id", $"{{{SetupID.ToString().ToUpper()}}}");
            AlternativeText = "Microsoft Office Signature Line...";
            ShowSignDate = true;
            AllowComments = false;
            SigningInstructions = "Before signing this document, verify that the content you are signing is correct.";
            _ws = ws;
            styleString = GetXmlNodeString(StyleAttrPath);
        }

        internal ExcelVmlDrawingSignatureLine(XmlNode topNode, XmlNamespaceManager ns, ExcelWorksheet ws) : base(topNode, ns)
        {
            var idString = GetXmlNodeString("o:signatureline/@id");
            SetupID = new Guid(idString);
            _ws = ws;
            styleString = GetXmlNodeString(StyleAttrPath);
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

        /// <summary>
        /// The ratio between EMU and Pixels
        /// </summary>
        public const int EMU_PER_PIXEL = 9525;
        /// <summary>
        /// The ratio between EMU and Points
        /// </summary>
        public const int EMU_PER_POINT = 12700;

        internal void GetToRowFromPixels(double pixels, out int toRow, out int rowOff, int fromRow = -1, int fromRowOff = -1)
        {
            if (fromRow < 0)
            {
                fromRow = From.Row;
                fromRowOff = From.RowOff;
            }
            ExcelWorksheet ws = _ws;
            var pixOff = pixels - ((ws.GetRowHeight(fromRow + 1) / 0.75) - (fromRowOff));
            double prevPixOff = pixels;
            int row = fromRow + 1;

            while (pixOff >= 0)
            {
                prevPixOff = pixOff;
                pixOff -= (ws.GetRowHeight(++row) / 0.75);
            }
            //Specific for signature lines?
            toRow = row - 1;
            if (fromRow == toRow)
            {
                rowOff = (int)(fromRowOff + (pixels));
            }
            else
            {
                rowOff = (int)(prevPixOff);
            }
        }

        internal void GetToColumnFromPixels(double pixels, out int col, out int colOff, int fromColumn = -1, int fromColumnOff = -1)
        {
            ExcelWorksheet ws = _ws;
            double mdw = ws.Workbook.MaxFontWidth;
            if (fromColumn < 0)
            {
                fromColumn = From.Column;
                fromColumnOff = From.ColumnOff;
            }
            double pixOff = pixels - (MathHelper.TruncateDouble(((256 * ws.GetColumnWidth(fromColumn + 1) + MathHelper.TruncateDouble(128 / mdw)) / 256) * mdw) - fromColumnOff);
            double offset = (double)fromColumnOff + pixels;
            col = fromColumn + 2;
            while (pixOff >= 0)
            {
                offset = pixOff;
                pixOff -= MathHelper.TruncateDouble(((256 * ws.GetColumnWidth(col++) + MathHelper.TruncateDouble(128 / mdw)) / 256) * mdw);
            }
            colOff = (int)offset;
        }

        internal void SetPixelHeight(double pixels)
        {
            GetToRowFromPixels(pixels, out int toRow, out int pixOff, From.Row, From.RowOffset);
            To.Row = toRow;
            To.RowOff = pixOff;

            UpdateXml();
        }

        internal void SetPixelWidth(double pixels)
        {
            GetToColumnFromPixels(pixels, out int col, out int pixOff);

            To.Column = col - 2;
            To.ColumnOff = pixOff;

            UpdateXml();
        }

        internal void UpdateXml()
        {
            From.UpdateXml();
            To.UpdateXml();

            var widthPt = (GetPixelWidth() * EMU_PER_PIXEL)/ EMU_PER_POINT;
            var heightPt = (GetPixelHeight() * EMU_PER_PIXEL)/ EMU_PER_POINT;
            SetStyle(styleString, "width", $"{widthPt}pt");
            SetStyle(styleString, "height", $"{heightPt}pt");
        }

        internal int pxWidth;
        internal int pxHeight;

        internal int GetPixelWidth()
        {
            var cols = From.Column - To.Column;
            double pix;
            double mdw = _ws.Workbook.MaxFontWidth;

            pix = -From.ColumnOff;
            for (int col = From.Column + 1; col <= To.Column; col++)
            {
                pix += MathHelper.TruncateDouble(((256 * _ws.GetColumnWidth(col) + MathHelper.TruncateDouble(128 / mdw)) / 256) * mdw);
            }

            var w = MathHelper.TruncateDouble(((256 * _ws.GetColumnWidth(To.Column + 1) + MathHelper.TruncateDouble(128 / mdw)) / 256) * mdw);
            pix += Math.Min(w, Convert.ToDouble(To.ColumnOff));

            return Convert.ToInt32(Math.Round(pix, MidpointRounding.AwayFromZero));
        }

        internal int GetPixelHeight()
        {
            ExcelWorksheet ws = _ws;

            double pix;

            pix = -(From.RowOff / (double)EMU_PER_PIXEL);
            for (int row = From.Row + 1; row <= To.Row; row++)
            {
                pix += ws.GetRowHeight(row) / 0.75;
            }
            var h = ws.GetRowHeight(To.Row + 1) / 0.75;
            pix += Math.Min(h, Convert.ToDouble(To.RowOff));

            return Convert.ToInt32(Math.Round(pix, MidpointRounding.AwayFromZero));
        }

        /// <summary>
        /// Set size in pixels
        /// Note that resizing columns / rows after using this function will effect the size of the drawing
        /// </summary>
        /// <param name="PixelWidth">Width in pixels</param>
        /// <param name="PixelHeight">Height in pixels</param>
        public void SetSize(int PixelWidth, int PixelHeight)
        {
            SetPixelWidth(PixelWidth);
            SetPixelHeight(PixelHeight);

            UpdateXml();
        }

        double _width = int.MinValue;
        double _height = int.MinValue;

        /// <summary>
        /// Set size in Percent.
        /// Note that resizing columns / rows after using this function will effect the size of the drawing
        /// </summary>
        /// <param name="Percent"></param>
        public void SetSize(int Percent)
        {
            if (_width == int.MinValue)
            {
                _width = GetPixelWidth();
                _height = GetPixelHeight();
            }
            _width = _width * ((double)Percent / 100);
            _height = _height * ((double)Percent / 100);

            SetPixelWidth(_width);
            SetPixelHeight(_height);

            UpdateXml();
        }
    }
}
