/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  11/18/2020         EPPlus Software AB       EPPlus 5.5
 *************************************************************************************************/
using OfficeOpenXml.Utils;
using OfficeOpenXml.Utils.Extensions;
using System;
using System.Xml;
namespace OfficeOpenXml.Drawing.Vml
{
    /// <summary>
    /// Border line settings for a vml drawing
    /// </summary>
    public class ExcelVmlDrawingBorder : XmlHelper
    {
        internal ExcelVmlDrawingBorder(ExcelDrawings drawings, XmlNamespaceManager ns, XmlNode topNode, string[] schemaNodeOrder) :
            base(ns, topNode)
        {
           SchemaNodeOrder = schemaNodeOrder;
        }
        /// <summary>
        /// The style of the border
        /// </summary>
        public eVmlLineStyle LineStyle 
        { 
            get
            {
                return GetXmlNodeString("v:stroke/@linestyle").ToEnum(eVmlLineStyle.NoLine);
            }
            set
            {
                if (value == eVmlLineStyle.NoLine)
                {
                    DeleteNode("v:stroke/@linestyle");
                    SetXmlNodeString("@stroked", "f");
                    DeleteNode("@strokeweight");
                }
                else
                {
                   SetXmlNodeString("v:stroke/@linestyle", value.ToEnumString());
                   DeleteNode("@stroked");
                }
            }
        }
        /// <summary>
        /// Dash style for the border 
        /// </summary>
        public eVmlDashStyle DashStyle 
        { 
            get
            {
                return CustomDashStyle.ToEnum(eVmlDashStyle.Custom);
            }
            set
            {
                CustomDashStyle = value.ToEnumString();
            }
        }
        /// <summary>
        /// Custom dash style.
        /// A series on numbers representing the width followed by the space between.        
        /// Example 1 : 8 2 1 2 1 2 --> Long dash dot dot. Space is twice the line width (2). LongDash (8) Dot (1). 
        /// Example 2 : 0 2 --> 0 implies a circular dot. 2 is the space between.
        /// </summary>
        public string CustomDashStyle
        {
            get
            {
                return GetXmlNodeString("v:stroke/@dashstyle");
            }
            set
            {
                SetXmlNodeString("v:stroke/@dashstyle", value);
            }
        }
        ExcelVmlMeasurementUnit _width = null;
        /// <summary>
        /// The width of the border
        /// </summary>
        public ExcelVmlMeasurementUnit Width
        {
            get
            {
                if(_width==null)
                {
                    _width = new ExcelVmlMeasurementUnit(GetXmlNodeString("@strokeweight"));
                }
                return _width;
            }
        }

        internal void UpdateXml()
        {
            if (_width != null)
            {
                if (Width.Value == 0)
                {
                    DeleteNode("@strokeweight");
                }
                else
                {
                    if (LineStyle == eVmlLineStyle.NoLine) LineStyle = eVmlLineStyle.Single;
                    SetXmlNodeString("@strokeweight", _width.GetValueString());
                }
            }
        }
    }
}
