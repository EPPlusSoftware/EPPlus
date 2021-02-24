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
using OfficeOpenXml.Drawing.Style;
using OfficeOpenXml.Drawing.Style.Fill;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Xml;

namespace OfficeOpenXml.Drawing.Theme
{
    /// <summary>
    /// Linestyle for a theme
    /// </summary>
    public class ExcelThemeLine : XmlHelper
    {
        internal ExcelThemeLine(XmlNamespaceManager nameSpaceManager, XmlNode topNode) : base(nameSpaceManager, topNode)
        {
            SchemaNodeOrder = new string[] { "noFill", "solidFill", "gradientFill", "pattFill", "prstDash", "round", "bevel", "miter", "headEnd", " tailEnd" };
        }
        const string widthPath = "@w";
        /// <summary>
        /// Line width, in EMU's
        /// 
        /// 1 Pixel      =   9525
        /// 1 Pt         =   12700
        /// 1 cm         =   360000 
        /// 1 US inch    =   914400
        /// </summary>
        public int Width
        {
            get
            {
                return GetXmlNodeInt(widthPath);
            }
            set
            {
                SetXmlNodeString(widthPath, value.ToString(CultureInfo.InvariantCulture));
            }
        }
        const string CapPath = "@cap";
        /// <summary>
        /// The ending caps for the line
        /// </summary>
        public eLineCap Cap
        {
            get
            {
                return EnumTransl.ToLineCap(GetXmlNodeString(CapPath));
            }
            set
            {
                SetXmlNodeString(CapPath, EnumTransl.FromLineCap(value));
            }
        }
        const string CompoundPath = "@cmpd";
        /// <summary>
        /// The compound line type to be used for the underline stroke
        /// </summary>
        public eCompundLineStyle CompoundLineStyle
        {
            get
            {
                return EnumTransl.ToLineCompound(GetXmlNodeString(CompoundPath));
            }
            set
            {
                SetXmlNodeString(CompoundPath, EnumTransl.FromLineCompound(value));
            }
        }
        const string PenAlignmentPath = "@algn";
        /// <summary>
        /// Specifies the pen alignment type for use within a text body
        /// </summary>
        public ePenAlignment Alignment
        {
            get
            {
                return EnumTransl.ToPenAlignment(GetXmlNodeString(PenAlignmentPath));
            }
            set
            {
                SetXmlNodeString(PenAlignmentPath, EnumTransl.FromPenAlignment(value));
            }
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
                    if (!(TopNode.HasChildNodes && TopNode.ChildNodes[0].LocalName.EndsWith("Fill")))
                    {
                        _fill = new ExcelDrawingFill(null, NameSpaceManager, TopNode.ChildNodes[0], "", SchemaNodeOrder);
                    }
                    else
                    {
                        var node = CreateNode("a:solidFill");
                        _fill = new ExcelDrawingFill(null, NameSpaceManager, TopNode.ChildNodes[0], "", SchemaNodeOrder);
                        Fill.SolidFill.Color.SetSchemeColor(eSchemeColor.Style);
                    }
                }
                return _fill;
            }
        }
        const string StylePath = "a:prstDash/@val";
        /// <summary>
        /// Preset line dash
        /// </summary>
        public eLineStyle Style
        {
            get
            {
                return EnumTransl.ToLineStyle(GetXmlNodeString(StylePath));
            }
            set
            {
                SetXmlNodeString(StylePath, EnumTransl.FromLineStyle(value));
            }
        }
        const string BevelPath = "a:bevel";
        const string RoundPath = "a:round";
        const string MiterPath = "a:miter";
        /// <summary>
        /// The shape that lines joined together have
        /// </summary>
        public eLineJoin? Join
        {
            get
            {
                if (ExistsNode(BevelPath))
                {
                    return eLineJoin.Bevel;
                }
                else if (ExistsNode(RoundPath))
                {
                    return eLineJoin.Round;
                }
                else if (ExistsNode(MiterPath))
                {
                    return eLineJoin.Miter;
                }
                else
                {
                    return null;
                }
            }
            set
            {
                if (value == eLineJoin.Bevel)
                {
                    CreateNode(BevelPath);
                    DeleteNode(RoundPath);
                    DeleteNode(MiterPath);
                }
                else if (value == eLineJoin.Round)
                {
                    CreateNode(RoundPath);
                    DeleteNode(BevelPath);
                    DeleteNode(MiterPath);
                }
                else
                {
                    CreateNode(MiterPath);
                    DeleteNode(RoundPath);
                    DeleteNode(BevelPath);
                }
            }
        }
        const string MiterJoinLimitPath = "a:miter/@lim";
        /// <summary>
        /// How much lines are extended to form a miter join
        /// </summary>
        public double? MiterJoinLimit
        {
            get
            {
                return GetXmlNodePercentage(MiterJoinLimitPath);
            }
            set
            {
                Join = eLineJoin.Miter;
                SetXmlNodePercentage(MiterJoinLimitPath, value);
            }
        }
        ExcelDrawingLineEnd _headEnd = null;
        /// <summary>
        /// Properties for drawing line head ends
        /// </summary>
        public ExcelDrawingLineEnd HeadEnd
        {
            get
            {
                if (_headEnd == null)
                {
                    return new ExcelDrawingLineEnd(NameSpaceManager, TopNode, "a:headEnd", Init);
                }
                return _headEnd;
            }
        }
        ExcelDrawingLineEnd _tailEnd = null;
        /// <summary>
        /// Properties for drawing line tail ends
        /// </summary>
        public ExcelDrawingLineEnd TailEnd
        {
            get
            {
                if (_tailEnd == null)
                {
                    return new ExcelDrawingLineEnd(NameSpaceManager, TopNode, "a:tailEnd", Init);
                }
                return _tailEnd;
            }
        }

        internal XmlElement LineElement
        {
            get
            {
                return TopNode as XmlElement;
            }
        }

        private void Init()
        {

        }
    }
}
