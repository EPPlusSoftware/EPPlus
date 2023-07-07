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
  07/07/2023         EPPlus Software AB       Epplus 7
 *************************************************************************************************/
using OfficeOpenXml.Utils.Extensions;
using System.Xml;

namespace OfficeOpenXml.ConditionalFormatting
{
    internal static class ColourScaleReadHandler
    {
        //We have no way of knowing what type of colorScale it is until we've read its first 3 nodes and xr is forward only.
        //This way we can determine type after reading the initial data.
        internal static ExcelConditionalFormattingRule CreateScales(ExcelAddress address, XmlReader xr, ExcelWorksheet ws)
        {
            //Read base rules
            var priority = int.Parse(xr.GetAttribute("priority"));
            var stopIfTrue = xr.GetAttribute("stopIfTrue") == "1";

            xr.Read();

            xr.Read();
            var lowType = xr.GetAttribute("type").ToEnum<eExcelConditionalFormattingValueObjectType>();
            string lowVal = xr.GetAttribute("val");

            if(address == null)
            {
                lowVal = ReadExtFormulaOrValue(xr, lowVal, lowType);
            }

            xr.Read();
            var middleOrHigh = xr.GetAttribute("type").ToEnum<eExcelConditionalFormattingValueObjectType>();
            string middleOrHighVal = xr.GetAttribute("val");

            if (address == null)
            {
                middleOrHighVal = ReadExtFormulaOrValue(xr, middleOrHighVal, middleOrHigh);
            }

            xr.Read();

            if (xr.LocalName == "color")
            {
                var twoColor = new ExcelConditionalFormattingTwoColorScale(
                    address, priority, ws, stopIfTrue, lowType, middleOrHigh, lowVal, middleOrHighVal, xr);

                twoColor.Type = eExcelConditionalFormattingRuleType.TwoColorScale;

                if (xr.LocalName == "cfRule" && xr.NodeType == XmlNodeType.EndElement)
                {
                    xr.Read();
                }

                return twoColor;
            }

            var highType = xr.GetAttribute("type").ToEnum<eExcelConditionalFormattingValueObjectType>();
            string highVal = xr.GetAttribute("val");

            if (address == null)
            {
                highVal = ReadExtFormulaOrValue(xr, highVal, highType);
            }

            xr.Read();

            var threeColor = new ExcelConditionalFormattingThreeColorScale(
                address, priority, ws, stopIfTrue, lowType, middleOrHigh, highType, lowVal, middleOrHighVal, highVal, xr);

            if (xr.LocalName == "cfRule" && xr.NodeType == XmlNodeType.EndElement)
            {
                xr.Read();
            }

            return threeColor;
        }

        private static string ReadExtFormulaOrValue(XmlReader xr, string valueNode, eExcelConditionalFormattingValueObjectType? type)
        {
            if (string.IsNullOrEmpty(valueNode) &&
                type != eExcelConditionalFormattingValueObjectType.Min &&
                type != eExcelConditionalFormattingValueObjectType.Max)
            {
                xr.Read();
                xr.Read();

                var content = xr.ReadContentAsString();

                xr.Read();

                return content;
            }

            return valueNode;
        }
    }
}
