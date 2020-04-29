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
using System.Xml;

namespace OfficeOpenXml.Drawing.Chart.ChartEx
{
    public class ExcelChartExStringData : ExcelChartExData
    {
        internal ExcelChartExStringData(XmlNamespaceManager nsm, XmlNode topNode) : base(nsm, topNode)
        {
        }
        public eStringDataType Type 
        {
            get
            {
                var s = GetXmlNodeString("@type");
                switch (s)
                {
                    case "entityId":
                        return eStringDataType.EntityId;
                    case "colorStr":
                        return eStringDataType.ColorString;
                    default:
                        return eStringDataType.Category;
                }
            }
            set
            {
                string s;
                switch (value)
                {
                    case eStringDataType.EntityId:
                        s = "entityId";
                        break;
                    case eStringDataType.ColorString:
                        s = "colorStr";
                        break;
                    default:
                        s = "cat";
                        break;
                }
                SetXmlNodeString("@type", s);
            }
        }
    }
}
