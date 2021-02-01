/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  04/29/2020         EPPlus Software AB           EPPlus 5.2
 *************************************************************************************************/
using OfficeOpenXml.Utils.Extensions;
using System.Xml;

namespace OfficeOpenXml.Drawing.Chart.ChartEx
{
    /// <summary>
    /// Numeric data reference for an extended chart
    /// </summary>
    public class ExcelChartExNumericData : ExcelChartExData
    {
        internal ExcelChartExNumericData(string worksheetName, XmlNamespaceManager nsm, XmlNode topNode) : base(worksheetName, nsm, topNode)
        {
        }
        /// <summary>
        /// The type of data.
        /// </summary>
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
