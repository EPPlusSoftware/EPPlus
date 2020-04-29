/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  04/29/2020         EPPlus Software AB       Initial release EPPlus 5
 *************************************************************************************************/
using OfficeOpenXml.Utils.Extentions;
using System.Xml;

namespace OfficeOpenXml.Drawing.Chart.ChartEx
{
    public class ExcelChartExNumericData : ExcelChartExData
    {
        internal ExcelChartExNumericData(XmlNamespaceManager nsm, XmlNode topNode) : base(nsm, topNode)
        {
        }
        public eNumericDataType Type 
        { 
            get
            {
                var s = GetXmlNodeString("@type");
                switch (s)
                {
                    case "val":
                        return eNumericDataType.Value;
                    case "colorVal":
                        return eNumericDataType.ColorValue;
                    default:
                        return s.ToEnum(eNumericDataType.Value);
                }
            }
            set
            {
                string s;
                switch(value)
                {
                    case eNumericDataType.Value:
                        s = "val";
                        break;
                    case eNumericDataType.ColorValue:
                        s = "colorVal";
                        break;
                    default:
                        s = value.ToEnumString();
                        break;
                }
                SetXmlNodeString("@type", s);
            }
        }
    }
}
