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
using System.Xml;
using OfficeOpenXml.Drawing.Interfaces;
using OfficeOpenXml.Drawing.Style.Effect;
using OfficeOpenXml.Drawing.Style.ThreeD;
using OfficeOpenXml.Style;

namespace OfficeOpenXml.Drawing.Chart
{
    /// <summary>
    /// Data table on chart level. 
    /// </summary>
    public class ExcelChartDataTable : XmlHelper, IDrawingStyle
    {
        ExcelChart _chart;
       internal ExcelChartDataTable(ExcelChart chart, XmlNamespaceManager ns, XmlNode node)
           : base(ns,node)
        {
            AddSchemaNodeOrder(new string[] { "dTable", "showHorzBorder", "showVertBorder", "showOutline", "showKeys", "spPr", "txPr" }, ExcelDrawing._schemaNodeOrderSpPr);
            XmlNode topNode = node.SelectSingleNode("c:dTable", NameSpaceManager);
           if (topNode == null)
           {
               topNode = node.OwnerDocument.CreateElement("c", "dTable", ExcelPackage.schemaChart);
               InserAfter(node, "c:valAx,c:catAx,c:dateAx,c:serAx", topNode);
               topNode.InnerXml = "<c:showHorzBorder val=\"1\"/><c:showVertBorder val=\"1\"/><c:showOutline val=\"1\"/><c:showKeys val=\"1\"/>" +
                    "<c:spPr><a:noFill/><a:ln cap = \"flat\" w=\"9525\" algn=\"ctr\" cmpd=\"sng\" ><a:solidFill><a:schemeClr val=\"tx1\"><a:lumMod val=\"15000\"/><a:lumOff val=\"85000\"/></a:schemeClr></a:solidFill><a:round/></a:ln><a:effectLst/></c:spPr>" +
                    "<c:txPr><a:bodyPr rot=\"0\" spcFirstLastPara=\"1\" vertOverflow=\"ellipsis\" vert=\"horz\" wrap=\"square\" anchor=\"ctr\" anchorCtr=\"1\"/>" +
                    "<a:lstStyle/><a:p><a:pPr rtl=\"0\"><a:defRPr sz=\"900\" b=\"0\" i=\"0\" u=\"none\" strike=\"noStrike\" kern=\"1200\" baseline=\"0\"><a:solidFill><a:schemeClr val=\"dk1\"></a:schemeClr></a:solidFill>" +
                    "<a:latin typeface=\" + mn - lt\"/><a:ea typeface=\" + mn - ea\"/><a:cs typeface=\" + mn - cs\"/></a:defRPr></a:pPr><a:endParaRPr lang=\"en - US\"/></a:p></c:txPr>";

           }
           TopNode = topNode;
            _chart = chart;
       }
       #region "Public properties"
       const string showHorzBorderPath = "c:showHorzBorder/@val";
        /// <summary>
        /// The horizontal borders will be shown in the data table
        /// </summary>
        public bool ShowHorizontalBorder
        {
           get
           {
               return GetXmlNodeBool(showHorzBorderPath);
           }
           set
           {
               SetXmlNodeString(showHorzBorderPath, value ? "1" : "0");
           }
       }
        const string showVertBorderPath = "c:showVertBorder/@val";
        /// <summary>
        /// The vertical borders will be shown in the data table
        /// </summary>
        public bool ShowVerticalBorder
        {
            get
            {
                return GetXmlNodeBool(showVertBorderPath);
            }
            set
            {
                SetXmlNodeString(showVertBorderPath, value ? "1" : "0");
            }
        }
        const string showOutlinePath = "c:showOutline/@val";
        /// <summary>
        /// The outline will be shown on the data table
        /// </summary>
        public bool ShowOutline
        {
            get
            {
                return GetXmlNodeBool(showOutlinePath);
            }
            set
            {
                SetXmlNodeString(showOutlinePath, value ? "1" : "0");
            }
        }
        const string showKeysPath = "c:showKeys/@val";
        /// <summary>
        /// The legend keys will be shown in the data table
        /// </summary>
        public bool ShowKeys
        {
            get
            {
                return GetXmlNodeBool(showKeysPath);
            }
            set
            {
                SetXmlNodeString(showKeysPath, value ? "1" : "0");
            }
        }
        ExcelDrawingFill _fill = null;
        /// <summary>
        /// Access fill properties
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
        /// Access border properties
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
        /// Access font properties
        /// </summary>
        public ExcelTextFont Font
        {
            get
            {
                if (_font == null)
                {
                    if (TopNode.SelectSingleNode("c:txPr", NameSpaceManager) == null)
                    {
                        CreateNode("c:txPr/a:bodyPr");
                        CreateNode("c:txPr/a:lstStyle");
                    }
                    _font = new ExcelTextFont(_chart, NameSpaceManager, TopNode, "c:txPr/a:p/a:pPr/a:defRPr", SchemaNodeOrder);
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
		ExcelDrawingTextSettings _textSettings = null;
		/// <summary>
		/// String settings like fills, text outlines and effects 
		/// </summary>
		public ExcelDrawingTextSettings TextSettings
		{
			get
			{
				if (_textSettings == null)
				{
					_textSettings = new ExcelDrawingTextSettings(_chart, NameSpaceManager, TopNode, $"c:txPr/a:p/a:pPr/a:defRPr", SchemaNodeOrder);
				}
				return _textSettings;
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

        #endregion
    }
}
