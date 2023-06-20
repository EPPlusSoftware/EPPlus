using OfficeOpenXml.Utils.Extensions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
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

            lowVal = ReadExtFormulaOrValue(xr, lowVal, lowType);

            xr.Read();
            var middleOrHigh = xr.GetAttribute("type").ToEnum<eExcelConditionalFormattingValueObjectType>();
            string middleOrHighVal = xr.GetAttribute("val");

            middleOrHighVal = ReadExtFormulaOrValue(xr, middleOrHighVal, middleOrHigh);

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

            highVal = ReadExtFormulaOrValue(xr, highVal, highType);

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
