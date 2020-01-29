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
using System.Globalization;
using System.Xml;
using OfficeOpenXml.Drawing.Interfaces;
using OfficeOpenXml.Drawing.Style.Effect;
using OfficeOpenXml.Drawing.Style.ThreeD;
using OfficeOpenXml.Style;

namespace OfficeOpenXml.Drawing.Chart
{
    /// <summary>
    /// A chart ledger
    /// </summary>
    public class ExcelChartLegend : XmlHelper, IDrawingStyle, IStyleMandatoryProperties
    {
        ExcelChart _chart;
        internal ExcelChartLegend(XmlNamespaceManager ns, XmlNode node, ExcelChart chart)
           : base(ns,node)
       {
           _chart=chart;
           AddSchemaNodeOrder(new string[] { "legendPos","legendEntry", "layout", "overlay", "spPr", "txPr" }, ExcelDrawing._schemaNodeOrderSpPr);
       }
        const string POSITION_PATH = "c:legendPos/@val";
        /// <summary>
        /// The position of the Legend
        /// </summary>
        public eLegendPosition Position 
        {
            get
            {
                switch(GetXmlNodeString(POSITION_PATH).ToLower(CultureInfo.InvariantCulture))
                {
                    case "t":
                        return eLegendPosition.Top;
                    case "b":
                        return eLegendPosition.Bottom;
                    case "l":
                        return eLegendPosition.Left;
                    case "tr":
                        return eLegendPosition.TopRight;
                    default:
                        return eLegendPosition.Right;
                }
            }
            set
            {
                if (TopNode == null) throw(new Exception("Can't set position. Chart has no legend"));
                switch (value)
                {
                    case eLegendPosition.Top:
                        SetXmlNodeString(POSITION_PATH, "t");
                        break;
                    case eLegendPosition.Bottom:
                        SetXmlNodeString(POSITION_PATH, "b");
                        break;
                    case eLegendPosition.Left:
                        SetXmlNodeString(POSITION_PATH, "l");
                        break;
                    case eLegendPosition.TopRight:
                        SetXmlNodeString(POSITION_PATH, "tr");
                        break;
                    default:
                        SetXmlNodeString(POSITION_PATH, "r");
                        break;
                }
            }
        }
        const string OVERLAY_PATH = "c:overlay/@val";
        /// <summary>
        /// If the legend overlays other objects
        /// </summary>
        public bool Overlay
        {
            get
            {
                return GetXmlNodeBool(OVERLAY_PATH);
            }
            set
            {
                if (TopNode == null) throw (new Exception("Can't set overlay. Chart has no legend"));
                SetXmlNodeBool(OVERLAY_PATH, value);
            }
        }
        ExcelDrawingFill _fill = null;
        /// <summary>
        /// The Fill style
        /// </summary>
        public ExcelDrawingFill Fill
        {
            get
            {
                if (_fill == null)
                {
                    _fill = new ExcelDrawingFill(_chart, NameSpaceManager, TopNode, "c:spPr", SchemaNodeOrder);
                }
                return _fill;
            }
        }
        ExcelDrawingBorder _border = null;
        /// <summary>
        /// The Border style
        /// </summary>
        public ExcelDrawingBorder Border
        {
            get
            {
                if (_border == null)
                {
                    _border = new ExcelDrawingBorder(_chart, NameSpaceManager, TopNode, "c:spPr/a:ln", SchemaNodeOrder);
                }
                return _border;
            }
        }
        ExcelTextFont _font = null;
        /// <summary>
        /// The Font properties
        /// </summary>
        public ExcelTextFont Font
        {
            get
            {
                if (_font == null)
                {

                    _font = new ExcelTextFont(_chart,NameSpaceManager, TopNode, "c:txPr/a:p/a:pPr/a:defRPr", SchemaNodeOrder);
                }
                return _font;
            }
        }
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
                    _textBody = new ExcelTextBody(NameSpaceManager, TopNode, "c:txPr/a:bodyPr", SchemaNodeOrder);
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
                    _effect = new ExcelDrawingEffectStyle(_chart, NameSpaceManager, TopNode, "c:spPr/a:effectLst", SchemaNodeOrder);
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
                    _threeD = new ExcelDrawing3D(NameSpaceManager, TopNode, "c:spPr", SchemaNodeOrder);
                }
                return _threeD;
            }
        }
        void IDrawingStyleBase.CreatespPr()
        {
            CreatespPrNode();
        }

        /// <summary>
        /// Remove the legend
        /// </summary>
        public void Remove()
        {
            if (TopNode == null) return;
            TopNode.ParentNode.RemoveChild(TopNode);
            TopNode = null;
        }
        /// <summary>
        /// Adds a legend to the chart
        /// </summary>
        public void Add()
        {
            if(TopNode!=null) return;

            //XmlHelper xml = new XmlHelper(NameSpaceManager, _chart.ChartXml);
            XmlHelper xml = XmlHelperFactory.Create(NameSpaceManager, _chart.ChartXml);
            xml.SchemaNodeOrder=_chart.SchemaNodeOrder;

            xml.CreateNode("c:chartSpace/c:chart/c:legend");
            TopNode = _chart.ChartXml.SelectSingleNode("c:chartSpace/c:chart/c:legend", NameSpaceManager);
            TopNode.InnerXml= "<c:legendPos val=\"r\" /><c:layout /><c:spPr><a:noFill/><a:ln><a:noFill/></a:ln><a:effectLst/></c:spPr>";
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

            CreatespPrNode();
        }
    }
}
