/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  01/27/2020         EPPlus Software AB       Initial release EPPlus 5
 *************************************************************************************************/
using OfficeOpenXml.Drawing.Style.Effect;
using OfficeOpenXml.Drawing.Style.ThreeD;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Xml;

namespace OfficeOpenXml.Drawing
{
    /// <summary>
    /// Base class for drawing-shape objects
    /// </summary>
    public class ExcelShapeBase : ExcelDrawing
    {
        internal string _shapeStylePath = "{0}xdr:spPr/a:prstGeom/@prst";
        private string _fillPath = "{0}xdr:spPr";
        private string _borderPath = "{0}xdr:spPr/a:ln";
        private string _effectPath = "{0}xdr:spPr/a:effectLst";
		private string _rotationPath = "{0}xdr:spPr/a:xfrm/@rot";
		private string _horizontalFlipPath = "{0}xdr:spPr/a:xfrm/@flipH";
		private string _verticalFlipPath = "{0}xdr:spPr/a:xfrm/@flipV";
		private string _headEndPath = "{0}xdr:spPr/a:ln/a:headEnd";
        private string _tailEndPath = "{0}xdr:spPr/a:ln/a:tailEnd";
        private string _textPath = "{0}xdr:txBody/a:p/a:r/a:t";
        private string _lockTextPath = "{0}@fLocksText";
        private string _textAnchoringPath = "{0}xdr:txBody/a:bodyPr/@anchor";
        private string _textAnchoringCtlPath = "{0}xdr:txBody/a:bodyPr/@anchorCtr";
        private string _paragraphPath = "{0}xdr:txBody/a:p";
        private string _textAlignPath = "{0}xdr:txBody/a:p/a:pPr/@algn";
        private string _indentAlignPath = "{0}xdr:txBody/a:p/a:pPr/@lvl";
        private string _textVerticalPath = "{0}xdr:txBody/a:bodyPr/@vert";
        private string _fontPath = "{0}xdr:txBody/a:p/a:pPr/a:defRPr";
        private string _textBodyPath = "{0}xdr:txBody/a:bodyPr";
        private string _presetGeometryPath = "{0}xdr:spPr/a:prstGeom/a:avLst";


        internal ExcelShapeBase(ExcelDrawings drawings, XmlNode node, string topPath, string nvPrPath, ExcelGroupShape parent=null) :
            base(drawings, node, topPath, nvPrPath, parent)
        {
            Init(string.IsNullOrEmpty(_topPath) ? "" : _topPath + "/");
        }
        private void Init(string topPath)
        {
            _shapeStylePath = string.Format(_shapeStylePath, topPath);
            _fillPath = string.Format(_fillPath, topPath);
            _borderPath = string.Format(_borderPath, topPath);
            _effectPath = string.Format(_effectPath, topPath);
			_rotationPath = string.Format(_rotationPath, topPath);
			_horizontalFlipPath = string.Format(_horizontalFlipPath, topPath);
			_verticalFlipPath = string.Format(_verticalFlipPath, topPath);
			_headEndPath = string.Format(_headEndPath, topPath);
            _tailEndPath = string.Format(_tailEndPath, topPath);
            _textPath = string.Format(_textPath, topPath);
            _lockTextPath = string.Format(_lockTextPath, topPath);
            _textAnchoringPath = string.Format(_textAnchoringPath, topPath);
            _textAnchoringCtlPath = string.Format(_textAnchoringCtlPath, topPath);
            _paragraphPath = string.Format(_paragraphPath, topPath);
            _textAlignPath = string.Format(_textAlignPath, topPath);
            _indentAlignPath = string.Format(_indentAlignPath, topPath);
            _textVerticalPath = string.Format(_textVerticalPath, topPath);
            _fontPath = string.Format(_fontPath, topPath);
            _textBodyPath = string.Format(_textBodyPath, topPath);
            _presetGeometryPath = string.Format(_presetGeometryPath, topPath);
            AddSchemaNodeOrder(SchemaNodeOrder, new string[] { "nvSpPr", "spPr", "txSp", "style", "txBody", "hlinkClick", "hlinkHover", "xfrm", "custGeom", "prstGeom", "noFill", "solidFill", "blipFill", "gradFill", "pattFill", "grpFill", "ln", "effectLst", "effectDag", "scene3d", "sp3d", "pPr", "r", "br", "fld", "endParaRPr", "lnRef", "fillRef", "effectRef", "fontRef" });
        }
        /// <summary>
        /// The type of drawing
        /// </summary>
        public override eDrawingType DrawingType
        {
            get
            {
                return eDrawingType.Shape;
            }
        }
        /// <summary>
        /// Shape style
        /// </summary>
        public virtual eShapeStyle Style
        {
            get
            {
                string v = GetXmlNodeString(_shapeStylePath);
                try
                {
                    return (eShapeStyle)Enum.Parse(typeof(eShapeStyle), v, true);
                }
                catch
                {
                    throw (new Exception(string.Format("Invalid shapetype {0}", v)));
                }
            }
            set
            {
                string v = value.ToString();
                v = v.Substring(0, 1).ToLower(CultureInfo.InvariantCulture) + v.Substring(1, v.Length - 1);
                SetXmlNodeString(_shapeStylePath, v);
            }
        }
        ExcelDrawingFill _fill = null;
        /// <summary>
        /// Access Fill properties
        /// </summary>
        public ExcelDrawingFill Fill
        {
            get
            {
                if (_fill == null)
                {
                    _fill = new ExcelDrawingFill(_drawings, NameSpaceManager, TopNode, _fillPath, SchemaNodeOrder);
                }
                return _fill;
            }
        }
        ExcelDrawingBorder _border = null;
        /// <summary>
        /// Access to Border propesties
        /// </summary>
        public ExcelDrawingBorder Border
        {
            get
            {
                if (_border == null)
                {
                    _border = new ExcelDrawingBorder(_drawings, NameSpaceManager, TopNode, _borderPath, SchemaNodeOrder);
                }
                return _border;
            }
        }
        ExcelDrawingEffectStyle _effect = null;
        /// <summary>
        /// Drawing effect properties
        /// </summary>
        public ExcelDrawingEffectStyle Effect
        {
            get
            {
                if (_effect == null)
                {
                    _effect = new ExcelDrawingEffectStyle(_drawings, NameSpaceManager, TopNode, _effectPath, SchemaNodeOrder);
                }
                return _effect;
            }
        }
        ExcelDrawing3D _threeD = null;
        /// <summary>
        /// Defines 3D properties to apply to an object
        /// </summary>
        public ExcelDrawing3D ThreeD
        {
            get
            {
                if (_threeD == null)
                {
                    _threeD = new ExcelDrawing3D(NameSpaceManager, TopNode, _fillPath, SchemaNodeOrder);
                }
                return _threeD;
            }
        }
        ExcelDrawingLineEnd _headEnd = null;
        /// <summary>
        /// Head line end
        /// </summary>
        public ExcelDrawingLineEnd HeadEnd
        {
            get
            {
                if (_headEnd == null)
                {
                    _headEnd = new ExcelDrawingLineEnd(NameSpaceManager, TopNode, _headEndPath, InitSpPr);
                }
                return _headEnd;
            }
        }
        ExcelDrawingLineEnd _tailEnd = null;
        /// <summary>
        /// Tail line end
        /// </summary>
        public ExcelDrawingLineEnd TailEnd
        {
            get
            {
                if (_tailEnd == null)
                {
                    _tailEnd = new ExcelDrawingLineEnd(NameSpaceManager, TopNode, _tailEndPath, InitSpPr);
                }
                return _tailEnd;
            }
        }
        ExcelTextFont _font = null;
        /// <summary>
        /// Font properties
        /// </summary>
        public ExcelTextFont Font
        {
            get
            {
                if (_font == null)
                {
                    XmlNode node = TopNode.SelectSingleNode(_paragraphPath, NameSpaceManager);
                    if (node == null)
                    {
                        Text = "";    //Creates the node p element
                        node = TopNode.SelectSingleNode(_paragraphPath, NameSpaceManager);
                    }
                    _font = new ExcelTextFont(_drawings, NameSpaceManager, TopNode, _fontPath, SchemaNodeOrder);
                }
                return _font;
            }
        }
        bool isSpInit = false;
        private void InitSpPr()
        {
            if (isSpInit == false)
            {
                if (CreateNodeUntil(_topPath, "spPr", out XmlNode spPrNode))
                {
                    spPrNode.InnerXml = "<a:noFill/><a:ln><a:noFill/></a:ln ><a:effectLst/><a:sp3d/>";
                }
            }
            isSpInit = true;
        }


        /// <summary>
        /// Text inside the shape
        /// </summary>
        public string Text
        {
            get
            {
                return RichText.Text;
            }
            set
            {                
                RichText.Text = value;
            }

        }
        /// <summary>
        /// Lock drawing
        /// </summary>
        public bool LockText
        {
            get
            {
                return GetXmlNodeBool(_lockTextPath, true);
            }
            set
            {
                SetXmlNodeBool(_lockTextPath, value);
            }
        }
        ExcelParagraphCollection _richText = null;
        internal static string[] _shapeNodeOrder= new string[] { "ln", "headEnd", "tailEnd", "effectLst", "blur", "fillOverlay", "glow", "innerShdw", "outerShdw", "prstShdw", "reflection", "softEdges", "effectDag", "scene3d", "scene3D", "sp3d", "bevelT", "bevelB", "extrusionClr", "contourClr" };

        /// <summary>
        /// Richtext collection. Used to format specific parts of the text
        /// </summary>
        public ExcelParagraphCollection RichText
        {
            get
            {
                if (_richText == null)
                {
                    _richText = new ExcelParagraphCollection(this, NameSpaceManager, TopNode, _paragraphPath, SchemaNodeOrder);
                }
                return _richText;
            }
        }
        /// <summary>
        /// Text Anchoring
        /// </summary>
        public eTextAnchoringType TextAnchoring
        {
            get
            {
                return GetXmlNodeString(_textAnchoringPath).TranslateTextAchoring();
            }
            set
            {
                SetXmlNodeString(_textAnchoringPath, value.TranslateTextAchoringText());
            }
        }
        /// <summary>
        /// The centering of the text box.
        /// </summary>
        public bool TextAnchoringControl
        {
            get
            {
                return GetXmlNodeBool(_textAnchoringCtlPath);
            }
            set
            {
                if (value)
                {
                    SetXmlNodeString(_textAnchoringCtlPath, "1");
                }
                else
                {
                    SetXmlNodeString(_textAnchoringCtlPath, "0");
                }
            }
        }
        /// <summary>
        /// How the text is aligned
        /// </summary>
        public eTextAlignment TextAlignment
        {
            get
            {
                switch (GetXmlNodeString(_textAlignPath))
                {
                    case "ctr":
                        return eTextAlignment.Center;
                    case "r":
                        return eTextAlignment.Right;
                    case "dist":
                        return eTextAlignment.Distributed;
                    case "just":
                        return eTextAlignment.Justified;
                    case "justLow":
                        return eTextAlignment.JustifiedLow;
                    case "thaiDist":
                        return eTextAlignment.ThaiDistributed;
                    default:
                        return eTextAlignment.Left;
                }
            }
            set
            {
                switch (value)
                {
                    case eTextAlignment.Right:
                        SetXmlNodeString(_textAlignPath, "r");
                        break;
                    case eTextAlignment.Center:
                        SetXmlNodeString(_textAlignPath, "ctr");
                        break;
                    case eTextAlignment.Distributed:
                        SetXmlNodeString(_textAlignPath, "dist");
                        break;
                    case eTextAlignment.Justified:
                        SetXmlNodeString(_textAlignPath, "just");
                        break;
                    case eTextAlignment.JustifiedLow:
                        SetXmlNodeString(_textAlignPath, "justLow");
                        break;
                    case eTextAlignment.ThaiDistributed:
                        SetXmlNodeString(_textAlignPath, "thaiDist");
                        break;
                    default:
                        DeleteNode(_textAlignPath);
                        break;
                }
            }
        }
        /// <summary>
        /// Indentation
        /// </summary>
        public int Indent
        {
            get
            {
                return GetXmlNodeInt(_indentAlignPath);
            }
            set
            {
                if (value < 0 || value > 8)
                {
                    throw (new ArgumentOutOfRangeException("Indent level must be between 0 and 8"));
                }
                SetXmlNodeString(_indentAlignPath, value.ToString());
            }
        }
		/// <summary>
		/// Rotation angle in degrees. Positive angles are clockwise. Negative angles are counter-clockwise.
		/// </summary>
		public double Rotation
        {
            get
            {
                return GetXmlNodeAngle(_rotationPath);
            }
            set
            {
                SetXmlNodeAngle(_rotationPath, value, "Rotation", -100000, 100000);
            }
        }
		/// <summary>
		/// If true, flips the shape horizontal about the center of its bounding box.
		/// </summary>
		public bool HorizontalFlip
        {
            get
            {
				return GetXmlNodeBool(_horizontalFlipPath);
			}
            set
            {
                SetXmlNodeBool(_horizontalFlipPath, value, false);
            }
        }
		/// <summary>
		/// If true, flips the shape vertical about the center of its bounding box.
		/// </summary>
		public bool VerticalFlip
		{
			get
			{
				return GetXmlNodeBool(_verticalFlipPath);
			}
			set
			{
				SetXmlNodeBool(_verticalFlipPath, value, false);
			}
		}
		/// <summary>
		/// Vertical text
		/// </summary>
		public eTextVerticalType TextVertical
        {
            get
            {
                return GetXmlNodeString(_textVerticalPath).TranslateTextVertical();
            }
            set
            {
                SetXmlNodeString(_textVerticalPath, value.TranslateTextVerticalText());
            }
        }
        ExcelTextBody _textBody = null;
        /// <summary>
        /// Access to text body properties.
        /// </summary>
        public ExcelTextBody TextBody
        {
            get
            {
                if (_textBody == null)
                {
                    _textBody = new ExcelTextBody(NameSpaceManager, TopNode, _textBodyPath, this.SchemaNodeOrder);
                }
                return _textBody;
            }
        }


        private static Dictionary<eShapeStyle, Dictionary<string, string>> ShapeGuide = new Dictionary<eShapeStyle, Dictionary<string, string>>()
        {
            { eShapeStyle.BentConnector3, Adjust1 },
            { eShapeStyle.CurvedConnector3, Adjust1 },
            { eShapeStyle.RoundRect, Adjust1 },
            { eShapeStyle.Snip1Rect, Adjust1 },
            { eShapeStyle.Snip2SameRect, Adjust2 },
            { eShapeStyle.Snip2DiagRect, Adjust2 },
            { eShapeStyle.SnipRoundRect, Adjust2 },
            { eShapeStyle.Round1Rect, Adjust1 },
            { eShapeStyle.Round2SameRect, Adjust2 },
            { eShapeStyle.Round2DiagRect, Adjust2 },
            { eShapeStyle.Triangle, Adjust1 },
            { eShapeStyle.Parallelogram, Adjust1 },
            { eShapeStyle.Trapezoid, Adjust1 },
            { eShapeStyle.Hexagon, AdjustVf },
            { eShapeStyle.Octagon, Adjust1 },
            { eShapeStyle.Pie, Adjust2 },
            { eShapeStyle.Chord, Adjust2 },
            { eShapeStyle.Teardrop, Adjust1 },
            { eShapeStyle.Frame, Adjust1 },
            { eShapeStyle.HalfFrame, Adjust2 },
            { eShapeStyle.Corner, Adjust2 },
            { eShapeStyle.DiagStripe, Adjust1 },
            { eShapeStyle.Plus, Adjust1 },
            { eShapeStyle.Plaque, Adjust1 },
            { eShapeStyle.Can, Adjust1 },
            { eShapeStyle.Cube, Adjust1 },
            { eShapeStyle.Bevel, Adjust1 },
            { eShapeStyle.Donut, Adjust1 },
            { eShapeStyle.NoSmoking, Adjust1 },
            { eShapeStyle.BlockArc, Adjust3 },
            { eShapeStyle.FoldedCorner¨, Adjust1 },
            { eShapeStyle.SmileyFace, Adjust1 },
            { eShapeStyle.Sun, Adjust1 },
            { eShapeStyle.Moon, Adjust1 },
            { eShapeStyle.Arc, Adjust2 },
            { eShapeStyle.BracketPair, Adjust1 },
            { eShapeStyle.BracePair, Adjust1 },
            { eShapeStyle.LeftBracket, Adjust1 },
            { eShapeStyle.RightBracket, Adjust1 },
            { eShapeStyle.LeftBrace, Adjust2 },
            { eShapeStyle.RightBrace, Adjust2 },
            { eShapeStyle.RightArrow, Adjust2 },
            { eShapeStyle.LeftArrow, Adjust2 },
            { eShapeStyle.UpArrow, Adjust2 },
            { eShapeStyle.DownArrow, Adjust2 },
            { eShapeStyle.LeftRightArrow, Adjust2 },
            { eShapeStyle.UpDownArrow, Adjust2 },
            { eShapeStyle.QuadArrow, Adjust3 },
            { eShapeStyle.LeftRightUpArrow, Adjust3 },
            { eShapeStyle.BentArrow, Adjust4 },
            { eShapeStyle.UturnArrow, Adjust5 },
            { eShapeStyle.LeftUpArrow, Adjust3 },
            { eShapeStyle.BentUpArrow, Adjust3 },
            { eShapeStyle.CurvedRightArrow, Adjust3 },
            { eShapeStyle.CurvedLeftArrow, Adjust3 },
            { eShapeStyle.CurvedUpArrow, Adjust3 },
            { eShapeStyle.CurvedDownArrow, Adjust3 },
            { eShapeStyle.StripedRightArrow, Adjust2 },
            { eShapeStyle.NotchedRightArrow, Adjust2 },
            { eShapeStyle.HomePlate, Adjust1 },
            { eShapeStyle.Chevron, Adjust1 },
            { eShapeStyle.RightArrowCallout, Adjust4 },
            { eShapeStyle.DownArrowCallout, Adjust4 },
            { eShapeStyle.LeftArrowCallout, Adjust4 },
            { eShapeStyle.UpArrowCallout, Adjust4 },
            { eShapeStyle.LeftRightArrowCallout, Adjust4 },
            { eShapeStyle.QuadArrowCallout, Adjust4 },
            { eShapeStyle.CircularArrow, Adjust5 },
            { eShapeStyle.MathPlus, Adjust1 },
            { eShapeStyle.MathMinus, Adjust1 },
            { eShapeStyle.MathMultiply, Adjust1 },
            { eShapeStyle.MathDivide, Adjust3 },
            { eShapeStyle.MathEqual, Adjust2 },
            { eShapeStyle.MathNotEqual, Adjust3 },
            { eShapeStyle.Star4, Adjust1 },
            { eShapeStyle.Star5, AdjustHfVf },
            { eShapeStyle.Star6, AdjustHf },
            { eShapeStyle.Star7, AdjustHfVf },
            { eShapeStyle.Star8, Adjust1 },
            { eShapeStyle.Star10, AdjustHf },
            { eShapeStyle.Star12, Adjust1 },
            { eShapeStyle.Star16, Adjust1 },
            { eShapeStyle.Star24, Adjust1 },
            { eShapeStyle.Star32, Adjust1 },
            { eShapeStyle.Ribbon2, Adjust2 },
            { eShapeStyle.Ribbon, Adjust2 },
            { eShapeStyle.EllipseRibbon2, Adjust3 },
            { eShapeStyle.EllipseRibbon, Adjust3 },
            { eShapeStyle.VerticalScroll, Adjust1 },
            { eShapeStyle.HorizontalScroll, Adjust1 },
            { eShapeStyle.Wave, Adjust2 },
            { eShapeStyle.DoubleWave, Adjust2 },
            { eShapeStyle.WedgeRectCallout, Adjust2 },
            { eShapeStyle.WedgeRoundRectCallout, Adjust3 },
            { eShapeStyle.WedgeEllipseCallout, Adjust2 },
            { eShapeStyle.CloudCallout, Adjust2 },
            { eShapeStyle.BorderCallout1, Adjust4 },
            { eShapeStyle.BorderCallout2, Adjust6 },
            { eShapeStyle.BorderCallout3, Adjust8 },
            { eShapeStyle.AccentCallout1, Adjust4 },
            { eShapeStyle.AccentCallout2, Adjust6 },
            { eShapeStyle.AccentCallout3, Adjust8 },
            { eShapeStyle.Callout1, Adjust4 },
            { eShapeStyle.Callout2, Adjust6 },
            { eShapeStyle.Callout3, Adjust8 },
            { eShapeStyle.AccentBorderCallout1, Adjust4 },
            { eShapeStyle.AccentBorderCallout2, Adjust6 },
            { eShapeStyle.AccentBorderCallout3, Adjust8 },
            { eShapeStyle.BentConnector3, Adjust1 },
            { eShapeStyle.CurvedConnector3, Adjust1 },
        };

        private static Dictionary<string, string> Adjust1 = new Dictionary<String, string>() { { "adj", "val " } };
        private static Dictionary<string, string> Adjust2 = new Dictionary<String, string>() { { "adj1", "val " }, {"adj2", "val " } };
        private static Dictionary<string, string> Adjust3 = new Dictionary<String, string>() { { "adj1", "val " }, { "adj2", "val " } , { "adj3", "val " } };
        private static Dictionary<string, string> Adjust4 = new Dictionary<String, string>() { { "adj1", "val " }, { "adj2", "val " }, { "adj3", "val " }, { "adj4", "val " } };
        private static Dictionary<string, string> Adjust5 = new Dictionary<String, string>() { { "adj1", "val " }, { "adj2", "val " }, { "adj3", "val " }, { "adj4", "val " }, { "adj5", "val " } };
        private static Dictionary<string, string> Adjust6 = new Dictionary<String, string>() { { "adj1", "val " }, { "adj2", "val " }, { "adj3", "val " }, { "adj4", "val " }, { "adj5", "val " }, { "adj6", "val " } };
        private static Dictionary<string, string> Adjust8 = new Dictionary<String, string>() { { "adj1", "val " }, { "adj2", "val " }, { "adj3", "val " }, { "adj4", "val " }, { "adj5", "val " }, { "adj6", "val " }, { "adj7", "val " }, { "adj8", "val " } };
        private static Dictionary<string, string> AdjustVf = new Dictionary<string, string>() { { "adj", "val " }, { "vf", "val " } };
        private static Dictionary<string, string> AdjustHf = new Dictionary<string, string>() { { "adj", "val " }, { "hf", "val " } };
        private static Dictionary<string, string> AdjustHfVf = new Dictionary<string, string>() { { "adj", "val " }, { "hf", "val " }, { "vf", "val " } };

        /* Set default values for each shape
         * when calling adjust shape we create a new dictionary and pull a copy of the one we want and populate with default values
         * change value we want to change
         * dict can't be added to or elements can't be removed once inited
         * ¨get list of keys in dict to change
         * several methods for edit that accept name and value, dictionary, params
         * 
         */

        //rename dictionary this should get layout from static dicts.
        private Dictionary <string, string> fmla = new Dictionary<string, string>();

        //returns a list of available shape adjustments keys
        public string[] GetShapeGuideAdjustments()
        {
            
        }

        //THis is THE main method, need new name? 
        public void EditShapeGuide(string name, int value)
        {
            if (fmla.ContainsKey(name)) throw new Exception("Name already exsists in Preset Geometry Nodes List");
            fmla[name] = value.ToString();
            //create xml node
            var gd = TopNode.SelectSingleNode(_presetGeometryPath + "/a:gd[@name =\"{name}\"]", NameSpaceManager);
            gd.Attributes["fmla"].Value = value.ToString();
        }

        //change to remove all 
        public void RemoveShapeGuideAdjustments()
        {
            fmla.Remove(name);
            var gd = TopNode.SelectSingleNode(_presetGeometryPath + "a:gd[@name =\"{name}\"]", NameSpaceManager);
            var parent = gd.ParentNode;
            parent.RemoveChild(gd);
        }


        internal override void CellAnchorChanged()
        {
            base.CellAnchorChanged();
            if (_fill != null) _fill.SetTopNode(TopNode);
            if (_border != null) _border.TopNode = TopNode;
            if (_effect != null) _effect.TopNode = TopNode;
            if (_font != null) _font.TopNode = TopNode;
            if (_threeD != null) _threeD.TopNode = TopNode;
            if (_tailEnd != null) _tailEnd.TopNode = TopNode;
            if (_headEnd != null) _headEnd.TopNode = TopNode;
            if (_richText != null) _richText.TopNode = TopNode;
            if (_textBody != null) _textBody.TopNode = TopNode;
        }
    }
}
