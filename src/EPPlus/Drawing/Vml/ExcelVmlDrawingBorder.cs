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
using OfficeOpenXml.Utils.Extentions;
using System;
using System.Xml;
namespace OfficeOpenXml.Drawing.Vml
{
    public class ExcelVmlDrawingBorder : XmlHelper
    {
        internal ExcelVmlDrawingBorder(ExcelDrawings drawings, XmlNamespaceManager ns, XmlNode topNode, string[] schemaNodeOrder) :
            base(ns, topNode)
        {
           SchemaNodeOrder = schemaNodeOrder;
        }
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
