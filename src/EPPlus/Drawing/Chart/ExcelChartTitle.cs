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
using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;
using OfficeOpenXml.Drawing.Chart.Style;
using OfficeOpenXml.Drawing.Interfaces;
using OfficeOpenXml.Drawing.Style.Effect;
using OfficeOpenXml.Drawing.Style.ThreeD;
using OfficeOpenXml.Style;

namespace OfficeOpenXml.Drawing.Chart
{
    /// <summary>
    /// The title of a chart
    /// </summary>
    public abstract class ExcelChartTitle : XmlHelper, IDrawingStyle, IStyleMandatoryProperties
    {
        internal ExcelChart _chart;
        internal string _nsPrefix = "";
        internal string _fontPropertiesPath = "";

        internal ExcelChartTitle(ExcelChart chart, XmlNamespaceManager nameSpaceManager, XmlNode node, string nsPrefix) :
            base(nameSpaceManager, node)
        {
            _chart = chart;
            _nsPrefix = nsPrefix;
            _fontPropertiesPath = $"{_nsPrefix}:tx/{_nsPrefix}:rich";
            if (chart._isChartEx)
            {
                AddSchemaNodeOrder(new string[] { "tx", "strRef", "rich", "bodyPr", "lstStyle", "layout", "p", "overlay", "spPr", "txPr" }, ExcelDrawing._schemaNodeOrderSpPr);
                CreateTopNode();
            }
            else
            {
                AddSchemaNodeOrder(_chart._chartXmlHelper.SchemaNodeOrder, ExcelDrawing._schemaNodeOrderSpPr);
                CreateTopNode();
                if (TopNode.HasChildNodes == false)
                {
                    TopNode.InnerXml = GetInitXml("c");
                    chart.ApplyStyleOnPart(this, chart.StyleManager?.Style?.Title, true);
                }
            }

        }

        private void CreateTopNode()
        {
            if (TopNode.LocalName != "title")
            {
                TopNode = CreateNode(_nsPrefix + ":title");
            }
        }

        internal static string GetInitXml(string prefix)
        {
            return $"<{prefix}:tx><{prefix}:rich><a:bodyPr rot=\"0\" spcFirstLastPara=\"1\" vertOverflow=\"ellipsis\" vert=\"horz\" wrap=\"square\" anchor=\"ctr\" anchorCtr=\"1\" />" +
                    $"<a:lstStyle />" +
                    $"<a:p><a:pPr>" +
                    $"<a:defRPr sz=\"1080\" b=\"1\" i=\"0\" u=\"none\" strike=\"noStrike\" kern=\"1200\" baseline=\"0\">" +
                    "<a:effectLst/><a:latin typeface=\"+mn-lt\"/><a:ea typeface=\"+mn-ea\"/><a:cs typeface=\"+mn-cs\"/></a:defRPr>" +
                    $"</a:pPr><a:r><a:t/></a:r></a:p></{prefix}:rich></{prefix}:tx><{prefix}:layout /><{prefix}:overlay val=\"0\" />" +
                    $"<{prefix}:spPr><a:noFill/><a:ln><a:noFill/></a:ln><a:effectLst/></{prefix}:spPr>";
        }

        /// <summary>
        /// The text
        /// </summary>
        public abstract string Text
        {
            get;
            set;
        }
        ExcelDrawingBorder _border = null;
        /// <summary>
        /// A reference to the border properties
        /// </summary>
        public ExcelDrawingBorder Border
        {
            get
            {
                if (_border == null)
                {
                    _border = new ExcelDrawingBorder(_chart, NameSpaceManager, TopNode, $"{_nsPrefix}:spPr/a:ln", SchemaNodeOrder);
                }
                return _border;
            }
        }
        ExcelDrawingFill _fill = null;
        /// <summary>
        /// A reference to the fill properties
        /// </summary>
        public ExcelDrawingFill Fill
        {
            get
            {
                if (_fill == null)
                {

                    _fill = new ExcelDrawingFill(_chart, NameSpaceManager, TopNode, $"{_nsPrefix}:spPr", SchemaNodeOrder);
                }
                return _fill;
            }
        }
        internal ExcelTextFont _font = null;
        /// <summary>
        /// A reference to the font properties
        /// </summary>
        public ExcelTextFont Font
        {
            get
            {
                if (_font == null)
                {
                    if (HasLinkedCell==false && (_richText == null || _richText.Count == 0))
                    {
                        RichText.Add("");
                    }
                    _font = new ExcelTextFont(_chart, NameSpaceManager, TopNode, $"{_fontPropertiesPath}/a:p/a:pPr/a:defRPr", SchemaNodeOrder);
                }
                return _font;
            }
        }
        internal abstract bool HasLinkedCell { get; }

        ExcelTextBody _textBody = null;
        /// <summary>
        /// Access to text body properties
        /// </summary>
        public ExcelTextBody TextBody
        {
            get
            {
                if (_textBody == null)
                {
                    _textBody = new ExcelTextBody(NameSpaceManager, TopNode, $"{_fontPropertiesPath}/a:bodyPr", SchemaNodeOrder);
                }
                return _textBody;
            }
        }
        ExcelDrawingEffectStyle _effect = null;
        /// <summary>
        /// Effects
        /// </summary>
        public ExcelDrawingEffectStyle Effect
        {
            get
            {
                if (_effect == null)
                {
                    _effect = new ExcelDrawingEffectStyle(_chart, NameSpaceManager, TopNode, $"{_nsPrefix}:spPr/a:effectLst", SchemaNodeOrder);
                }
                return _effect;
            }
        }
        ExcelDrawing3D _threeD = null;
        /// <summary>
        /// 3D properties
        /// </summary>
        public ExcelDrawing3D ThreeD
        {
            get
            {
                if (_threeD == null)
                {
                    _threeD = new ExcelDrawing3D(NameSpaceManager, TopNode, $"{_nsPrefix}:spPr", SchemaNodeOrder);
                }
                return _threeD;
            }
        }
        void IDrawingStyleBase.CreatespPr()
        {
            CreatespPrNode($"{_nsPrefix}:spPr");
        }

        ExcelParagraphCollection _richText = null;
        /// <summary>
        /// Richtext. If the title is linked to a cell via <see cref="ExcelChartTitleStandard.LinkedCell"/>, this property returns null
        /// </summary>
        public ExcelParagraphCollection RichText
        {
            get
            {
                if (HasLinkedCell) return null;
                if (_richText == null)
                {
                    CreateRichText();
                }
                return _richText;
            }
        }

        internal void CreateRichText()
        {
            float defFont = 14;
            var stylePart = GetStylePart();
            if (stylePart != null && stylePart.HasTextRun)
            {
                defFont = Convert.ToSingle(stylePart.DefaultTextRun.FontSize);
            }
            _richText = new ExcelParagraphCollection(_chart, NameSpaceManager, TopNode, $"{_fontPropertiesPath}/a:p", SchemaNodeOrder, defFont);
        }

        private ExcelChartStyleEntry GetStylePart()
        {
            var style = _chart._styleManager?.Style;
            if (style == null) return null;
            if (TopNode.ParentNode.LocalName == "chart")
            {
                return _chart._styleManager.Style.Title;
            }
            else
            {
                return _chart._styleManager.Style.AxisTitle;
            }
        }

        /// <summary>
        /// Show without overlaping the chart.
        /// </summary>
        public bool Overlay
        {
            get
            {
                if (_chart._isChartEx)
                {
                    return GetXmlNodeBool("@overlay");
                }
                else
                {
                    return GetXmlNodeBool("c:overlay/@val");
                }
            }
            set
            {
                if (_chart._isChartEx)
                {
                    SetXmlNodeBool("@overlay", value);
                }
                else
                {
                    SetXmlNodeBool("c:overlay/@val", value);
                }
            }
        }
        /// <summary>
        /// The centering of the text. Centers the text to the smallest possible text container.
        /// </summary>
        public bool AnchorCtr
        {
            get
            {
                return GetXmlNodeBool($"{_fontPropertiesPath}/a:bodyPr/@anchorCtr", false);
            }
            set
            {
                SetXmlNodeBool($"{_fontPropertiesPath}/a:bodyPr/@anchorCtr", value, false);
            }
        }
        /// <summary>
        /// How the text is anchored
        /// </summary>
        public eTextAnchoringType Anchor
        {
            get
            {
                return GetXmlNodeString($"{_fontPropertiesPath}/a:bodyPr/@anchor").TranslateTextAchoring();
            }
            set
            {
                SetXmlNodeString($"{_fontPropertiesPath}/a:bodyPr/@anchorCtr", value.TranslateTextAchoringText());
            }
        }
        const string TextVerticalPath = "xdr:sp/xdr:txBody/a:bodyPr/@vert";
        /// <summary>
        /// Vertical text
        /// </summary>
        public eTextVerticalType TextVertical
        {
            get
            {
                return GetXmlNodeString($"{_fontPropertiesPath}/a:bodyPr/@vert").TranslateTextVertical();
            }
            set
            {
                SetXmlNodeString($"{_fontPropertiesPath}/a:bodyPr/@vert", value.TranslateTextVerticalText());
            }
        }
        /// <summary>
        /// Rotation in degrees (0-360)
        /// </summary>
        public double Rotation
        {
            get
            {
                var i=GetXmlNodeInt($"{_fontPropertiesPath}/a:bodyPr/@rot");
                if (i < 0)
                {
                    return 360 - (i / 60000);
                }
                else
                {
                    return (i / 60000);
                }
            }
            set
            {
                int v;
                if(value <0 || value > 360)
                {
                    throw(new ArgumentOutOfRangeException("Rotation must be between 0 and 360"));
                }

                if (value > 180)
                {
                    v = (int)((value - 360) * 60000);
                }
                else
                {
                    v = (int)(value * 60000);
                }
                SetXmlNodeString($"{_fontPropertiesPath}/a:bodyPr/@rot", v.ToString());
            }
        }

        void IStyleMandatoryProperties.SetMandatoryProperties()
        {
            TextBody.Anchor = eTextAnchoringType.Center;
            TextBody.AnchorCenter = true;
            TextBody.WrapText = eTextWrappingType.Square;
            TextBody.VerticalTextOverflow = eTextVerticalOverflow.Ellipsis;
            TextBody.ParagraphSpacing = true;
            TextBody.Rotation = 0;

            if (Font.Kerning == 0) Font.Kerning = 12;
            Font.Bold = Font.Bold; //Must be set

            CreatespPrNode($"{_nsPrefix}:spPr");
        }
    }
    /// <summary>
    /// A chart title for a standard chart.
    /// </summary>
    public class ExcelChartTitleStandard : ExcelChartTitle
    {
        internal ExcelChartTitleStandard(ExcelChart chart, XmlNamespaceManager nameSpaceManager, XmlNode node, string nsPrefix) : base(chart, nameSpaceManager, node, nsPrefix)
        {
        }
        public override string Text 
        {
            get
            {
                if (LinkedCell == null)
                {
                    return RichText.Text;
                }
                else
                {
                    return LinkedCell.Text;
                }
            }
            set
            {
                if(RichText == null)
                {
                    LinkedCell = null;
                    CreateRichText();
                }
                var applyStyle = (RichText.Count == 0);
                RichText.Text = value;
                _font = null;
                if (applyStyle) _chart.ApplyStyleOnPart(this, _chart.StyleManager?.Style?.Title, true);
            }
        }
        /// <summary>
        /// A reference to a cell used as the title text
        /// </summary>
        public ExcelRangeBase LinkedCell
        {
            get
            {
                var a = GetXmlNodeString($"{_nsPrefix}:tx/{_nsPrefix}:strRef/c:f");
                if (ExcelCellBase.IsValidAddress(a))
                {
                    var address = new ExcelAddressBase(a);
                    ExcelWorksheet ws;
                    if (string.IsNullOrEmpty(address.WorkSheetName))
                    {
                        ws = _chart.WorkSheet;
                    }
                    else
                    {
                        ws = _chart.WorkSheet.Workbook.Worksheets[address.WorkSheetName];
                    }
                    if (ws == null) return null;
                    return ws.Cells[address.LocalAddress];
                }
                return null;
            }
            set
            {
                if (value == null)
                {
                    DeleteNode($"{_nsPrefix}:tx/{_nsPrefix}:strRef");
                    RichText.Text = "";
                    _fontPropertiesPath = $"{_nsPrefix}:tx/{_nsPrefix}:rich";                    
                }
                else
                {
                    DeleteNode($"{_nsPrefix}:tx/{_nsPrefix}:rich");
                    SetXmlNodeString($"{_nsPrefix}:tx/{_nsPrefix}:strRef/c:f", value.FullAddressAbsolute);
                    var cache = CreateNode($"{_nsPrefix}:tx/{_nsPrefix}:strRef/c:strCache", false, true);
                    cache.InnerXml = $"<c:ptCount val=\"1\"/><c:pt idx=\"0\"><c:v>{value.Text}</c:v></c:pt>";
                    _fontPropertiesPath = $"{_nsPrefix}:txPr";
                }
                _font = null;
            }
        }
        internal override bool HasLinkedCell => LinkedCell != null;
    }
}
