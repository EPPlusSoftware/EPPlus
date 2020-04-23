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
using OfficeOpenXml.Drawing.Interfaces;
using OfficeOpenXml.Drawing.Style.Effect;
using OfficeOpenXml.Drawing.Style.ThreeD;
using OfficeOpenXml.Style;
using System.Xml;

namespace OfficeOpenXml.Drawing.Chart
{
    /// <summary>
    /// Access to trendline label properties
    /// </summary>
    public class ExcelChartTrendlineLabel : XmlHelper, IDrawingStyle
    {
        ExcelChartSerieStandard _serie;        
        internal ExcelChartTrendlineLabel(XmlNamespaceManager namespaceManager, XmlNode topNode, ExcelChartSerieStandard serie) : base(namespaceManager, topNode)
        {
            _serie = serie;

            AddSchemaNodeOrder(new string[] { "layout", "tx", "numFmt", "spPr", "txPr" }, ExcelDrawing._schemaNodeOrderSpPr);
        }

        ExcelDrawingFill _fill = null;
        /// <summary>
        /// Access to fill properties
        /// </summary>
        public ExcelDrawingFill Fill
        {
            get
            {
                if (_fill == null)
                {
                    _fill = new ExcelDrawingFill(_serie._chart, NameSpaceManager, TopNode, "c:trendlineLbl/c:spPr", SchemaNodeOrder);
                }
                return _fill;
            }
        }
        ExcelDrawingBorder _border = null;
        /// <summary>
        /// Access to border properties
        /// </summary>
        public ExcelDrawingBorder Border
        {
            get
            {
                if (_border == null)
                {
                    _border = new ExcelDrawingBorder(_serie._chart, NameSpaceManager, TopNode, "c:trendlineLbl/c:spPr/a:ln", SchemaNodeOrder);
                }
                return _border;
            }
        }
        ExcelTextFont _font = null;
        /// <summary>
        /// Access to font properties
        /// </summary>
        public ExcelTextFont Font
        {
            get
            {
                if (_font == null)
                {
                    _font = new ExcelTextFont(_serie._chart, NameSpaceManager, TopNode, "c:trendlineLbl/c:txPr/a:p/a:pPr/a:defRPr", SchemaNodeOrder);
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
                    _textBody = new ExcelTextBody(NameSpaceManager, TopNode, "c:trendlineLbl/c:txPr/a:bodyPr", SchemaNodeOrder);
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
                    _effect = new ExcelDrawingEffectStyle(_serie._chart, NameSpaceManager, TopNode, "c:trendlineLbl/c:spPr/a:effectLst", SchemaNodeOrder);
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
                    _threeD = new ExcelDrawing3D(NameSpaceManager, TopNode, "c:trendlineLbl/c:spPr", SchemaNodeOrder);
                }
                return _threeD;
            }
        }
        void IDrawingStyleBase.CreatespPr()
        {
            CreatespPrNode("c:trendlineLbl/c:spPr");
        }

        ExcelParagraphCollection _richText = null;
        /// <summary>
        /// Richtext
        /// </summary>
        public ExcelParagraphCollection RichText
        {
            get
            {
                if (_richText == null)
                {
                    _richText = new ExcelParagraphCollection(_serie._chart, NameSpaceManager, TopNode, "c:trendlineLbl/c:tx/c:rich/a:p", SchemaNodeOrder);
                }
                return _richText;
            }
        }
        /// <summary>
        /// Numberformat
        /// </summary>
        public string NumberFormat
        {
            get
            {
                return GetXmlNodeString("c:trendlineLbl/c:numFmt/@formatCode");
            }
            set
            {
                SetXmlNodeString("c:trendlineLbl/c:numFmt/@formatCode", value);
            }
        }
        /// <summary>
        /// If the numberformat is linked to the source data
        /// </summary>
        public bool SourceLinked
        {
            get
            {
                return GetXmlNodeBool("c:trendlineLbl/c:numFmt/@sourceLinked");
            }
            set
            {
                SetXmlNodeBool("c:trendlineLbl/c:numFmt/@sourceLinked", value, true);
            }
        }        
    }
}