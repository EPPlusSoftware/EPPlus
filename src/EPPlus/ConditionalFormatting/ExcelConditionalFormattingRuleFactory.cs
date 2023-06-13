using OfficeOpenXml.ConditionalFormatting;
using OfficeOpenXml.ConditionalFormatting.Contracts;
using OfficeOpenXml.ConditionalFormatting.Rules;
using OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Text;
using OfficeOpenXml.Utils;
using OfficeOpenXml.Utils.Extensions;
using System;
using System.Runtime.CompilerServices;
using System.Xml;

namespace OfficeOpenXml.ConditionalFormatting
{
    internal static class ExcelConditionalFormattingRuleFactory
    {

        public static ExcelConditionalFormattingRule Create(
        eExcelConditionalFormattingRuleType type,
        ExcelAddress address,
        int priority, ExcelWorksheet worksheet)
        {
            Require.Argument(type);
            Require.Argument(worksheet).IsNotNull("worksheet");

            switch (type)
            {
                case eExcelConditionalFormattingRuleType.GreaterThan:

                    return new ExcelConditionalFormattingGreaterThan(
                          address,
                          priority,
                          worksheet);

                case eExcelConditionalFormattingRuleType.GreaterThanOrEqual:
                    return new ExcelConditionalFormattingGreaterThanOrEqual(
                          address,
                          priority,
                          worksheet);

                case eExcelConditionalFormattingRuleType.LessThan:
                    return new ExcelConditionalFormattingLessThan(
                       address,
                       priority,
                       worksheet);

                case eExcelConditionalFormattingRuleType.LessThanOrEqual:
                    return new ExcelConditionalFormattingLessThanOrEqual(
                          address,
                          priority,
                          worksheet);

                case eExcelConditionalFormattingRuleType.Between:
                    return new ExcelConditionalFormattingBetween(
                       address,
                       priority,
                       worksheet);

                case eExcelConditionalFormattingRuleType.NotBetween:
                    return new ExcelConditionalFormattingNotBetween(
                       address,
                       priority,
                       worksheet);

                case eExcelConditionalFormattingRuleType.Equal:
                    return new ExcelConditionalFormattingEqual(
                       address,
                       priority,
                       worksheet);

                case eExcelConditionalFormattingRuleType.NotEqual:
                    return new ExcelConditionalFormattingNotEqual(
                        address, 
                        priority, 
                        worksheet);

                case eExcelConditionalFormattingRuleType.ContainsText:
                    return new ExcelConditionalFormattingContainsText(
                       address,
                       priority,
                       worksheet);

                case eExcelConditionalFormattingRuleType.NotContainsText:
                    return new ExcelConditionalFormattingNotContainsText(
                       address,
                       priority,
                       worksheet);

                case eExcelConditionalFormattingRuleType.ContainsErrors:
                    return new ExcelConditionalFormattingContainsErrors(
                       address,
                       priority,
                       worksheet);

                case eExcelConditionalFormattingRuleType.NotContainsErrors:
                    return new ExcelConditionalFormattingNotContainsErrors(
                       address,
                       priority,
                       worksheet);

                case eExcelConditionalFormattingRuleType.BeginsWith:
                    return new ExcelConditionalFormattingBeginsWith(
                       address,
                       priority,
                       worksheet);

                case eExcelConditionalFormattingRuleType.EndsWith:
                    return new ExcelConditionalFormattingEndsWith(
                       address,
                       priority,
                       worksheet);

                case eExcelConditionalFormattingRuleType.ContainsBlanks:
                    return new ExcelConditionalFormattingContainsBlanks(
                       address,
                       priority,
                       worksheet);

                case eExcelConditionalFormattingRuleType.NotContainsBlanks:
                    return new ExcelConditionalFormattingNotContainsBlanks(
                       address,
                       priority,
                       worksheet);

                case eExcelConditionalFormattingRuleType.Expression:
                    return new ExcelConditionalFormattingExpression(
                        address, 
                        priority, 
                        worksheet);

                case eExcelConditionalFormattingRuleType.Yesterday:
                    return new ExcelConditionalFormattingYesterday(
                        address,
                        priority,
                        worksheet);

                case eExcelConditionalFormattingRuleType.Today:
                    return new ExcelConditionalFormattingToday(
                        address,
                        priority,
                        worksheet);

                case eExcelConditionalFormattingRuleType.Tomorrow:
                    return new ExcelConditionalFormattingTomorrow(
                        address,
                        priority,
                        worksheet);

                case eExcelConditionalFormattingRuleType.Last7Days:
                    return new ExcelConditionalFormattingLast7Days(
                        address,
                        priority,
                        worksheet);

                case eExcelConditionalFormattingRuleType.LastWeek:
                    return new ExcelConditionalFormattingLastWeek(
                        address,
                        priority,
                        worksheet);

                case eExcelConditionalFormattingRuleType.ThisWeek:
                    return new ExcelConditionalFormattingThisWeek(
                        address,
                        priority,
                        worksheet);

                case eExcelConditionalFormattingRuleType.NextWeek:
                    return new ExcelConditionalFormattingNextWeek(
                        address,
                        priority,
                        worksheet);

                case eExcelConditionalFormattingRuleType.LastMonth:
                    return new ExcelConditionalFormattingLastMonth(
                        address,
                        priority,
                        worksheet);

                case eExcelConditionalFormattingRuleType.ThisMonth:
                    return new ExcelConditionalFormattingThisMonth(
                        address,
                        priority,
                        worksheet);

                case eExcelConditionalFormattingRuleType.NextMonth:
                    return new ExcelConditionalFormattingNextMonth(
                        address,
                        priority,
                        worksheet);

                case eExcelConditionalFormattingRuleType.DuplicateValues:
                    return new ExcelConditionalFormattingDuplicateValues(
                        address,
                        priority,
                        worksheet);

                case eExcelConditionalFormattingRuleType.Top:
                case eExcelConditionalFormattingRuleType.Bottom:
                case eExcelConditionalFormattingRuleType.TopPercent:
                case eExcelConditionalFormattingRuleType.BottomPercent:
                    return new ExcelConditionalFormattingTopBottomGroup(
                        type, 
                        address,
                        priority,
                        worksheet);

                case eExcelConditionalFormattingRuleType.AboveAverage:
                case eExcelConditionalFormattingRuleType.AboveOrEqualAverage:
                case eExcelConditionalFormattingRuleType.BelowOrEqualAverage:
                case eExcelConditionalFormattingRuleType.BelowAverage:
                    return new ExcelConditionalFormattingAverageGroup(
                        type, 
                        address,
                        priority,
                        worksheet);

                case eExcelConditionalFormattingRuleType.AboveStdDev:
                case eExcelConditionalFormattingRuleType.BelowStdDev:
                    return new ExcelConditionalFormattingStdDevGroup(
                        type, 
                        address,
                        priority,
                        worksheet);

                case eExcelConditionalFormattingRuleType.UniqueValues:
                    return new ExcelConditionalFormattingUniqueValues(
                        address,
                        priority,
                        worksheet);

                case eExcelConditionalFormattingRuleType.DataBar:
                    return new ExcelConditionalFormattingDataBar(
                        address, 
                        priority,
                        worksheet);

                case eExcelConditionalFormattingRuleType.TwoColorScale:
                    return new ExcelConditionalFormattingTwoColorScale(
                        address,
                        priority,
                        worksheet);

                case eExcelConditionalFormattingRuleType.ThreeColorScale:
                    return new ExcelConditionalFormattingThreeColorScale(
                        address,
                        priority,
                        worksheet);

                case eExcelConditionalFormattingRuleType.ThreeIconSet:
                    return new ExcelConditionalFormattingThreeIconSet(
                        address,
                        priority,
                        worksheet);

                case eExcelConditionalFormattingRuleType.FourIconSet:
                    return new ExcelConditionalFormattingFourIconSet(
                        address,
                        priority,
                        worksheet);

                case eExcelConditionalFormattingRuleType.FiveIconSet:
                    return new ExcelConditionalFormattingFiveIconSet(
                        address,
                        priority,
                        worksheet);
            }

            throw new InvalidOperationException(
             string.Format(
             ExcelConditionalFormattingConstants.Errors.NonSupportedRuleType,
             type.ToString()));
        }

        public static ExcelConditionalFormattingRule Create(ExcelAddress address, ExcelWorksheet ws, XmlReader xr)
        {
            string cfType = xr.GetAttribute("type");
            string op = xr.GetAttribute("operator");

            if (cfType == "cellIs")
            {
                cfType = op;
            }

            if (cfType == "timePeriod")
            {
                cfType = xr.GetAttribute("timePeriod");
            }

            if (cfType == "top10")
            {
                bool isPercent = !string.IsNullOrEmpty(xr.GetAttribute("percent"));
                bool isBottom = !string.IsNullOrEmpty(xr.GetAttribute("bottom"));

                if (isPercent)
                {
                    cfType = "TopPercent";

                    if (isBottom)
                    {
                        cfType = "BottomPercent";
                    }
                }
                else if (isBottom)
                {
                    cfType = "Bottom";
                }
                else
                {
                    cfType = "Top";
                }
            }

            if (cfType == "aboveAverage")
            {
                //aboveAverage is true by default/when empty
                if (string.IsNullOrEmpty(xr.GetAttribute("aboveAverage")))
                {
                    cfType = "Above";
                }
                else
                {
                    cfType = "Below";
                }

                string stringEnding = "Average";

                if (!string.IsNullOrEmpty(xr.GetAttribute("stdDev")))
                {
                    stringEnding = "StdDev";
                }
                else if (!string.IsNullOrEmpty(xr.GetAttribute("equalAverage")))
                {
                    cfType = cfType + "OrEqual";
                }

                cfType = cfType + stringEnding;
            }

            string text = xr.GetAttribute("timePeriod");

            if(cfType == "colorScale")
            {
                return ColourScaleReadHandler.CreateScales(address, xr, ws);
            }

            if(cfType == "iconSet")
            {
                return IconReadHandler.ReadIcons(address, xr, ws);
            }

            var eType = cfType.ToEnum<eExcelConditionalFormattingRuleType>().Value;

            switch (eType)
            {
                case eExcelConditionalFormattingRuleType.GreaterThan:
                    return new ExcelConditionalFormattingGreaterThan(address, ws, xr);

                case eExcelConditionalFormattingRuleType.GreaterThanOrEqual:
                    return new ExcelConditionalFormattingGreaterThanOrEqual(address, ws, xr);

                case eExcelConditionalFormattingRuleType.LessThan:
                    return new ExcelConditionalFormattingLessThan(address, ws, xr);

                case eExcelConditionalFormattingRuleType.LessThanOrEqual:
                    return new ExcelConditionalFormattingLessThanOrEqual(address, ws, xr);

                case eExcelConditionalFormattingRuleType.Between:
                    return new ExcelConditionalFormattingBetween(address, ws, xr);

                case eExcelConditionalFormattingRuleType.NotBetween:
                    return new ExcelConditionalFormattingNotBetween(address, ws, xr);

                case eExcelConditionalFormattingRuleType.Equal:
                    return new ExcelConditionalFormattingEqual(address, ws, xr);

                case eExcelConditionalFormattingRuleType.NotEqual:
                    return new ExcelConditionalFormattingNotEqual(address, ws, xr);

                case eExcelConditionalFormattingRuleType.ContainsText:
                    return new ExcelConditionalFormattingContainsText(address, ws, xr);

                case eExcelConditionalFormattingRuleType.NotContainsText:
                    return new ExcelConditionalFormattingNotContainsText(address, ws, xr);

                case eExcelConditionalFormattingRuleType.BeginsWith:
                    return new ExcelConditionalFormattingBeginsWith(address, ws, xr);

                case eExcelConditionalFormattingRuleType.EndsWith:
                    return new ExcelConditionalFormattingEndsWith(address, ws, xr);

                case eExcelConditionalFormattingRuleType.Expression:
                    return new ExcelConditionalFormattingExpression(address, ws, xr);

                case eExcelConditionalFormattingRuleType.ContainsBlanks:
                    return new ExcelConditionalFormattingContainsBlanks(address, ws, xr);

                case eExcelConditionalFormattingRuleType.NotContainsBlanks:
                    return new ExcelConditionalFormattingNotContainsBlanks(address, ws, xr);

                case eExcelConditionalFormattingRuleType.ContainsErrors:
                    return new ExcelConditionalFormattingContainsErrors(address, ws, xr);

                case eExcelConditionalFormattingRuleType.NotContainsErrors:
                    return new ExcelConditionalFormattingNotContainsErrors(address, ws, xr);

                case eExcelConditionalFormattingRuleType.Yesterday:
                    return new ExcelConditionalFormattingYesterday(address, ws, xr);

                case eExcelConditionalFormattingRuleType.Today:
                    return new ExcelConditionalFormattingToday(address, ws, xr);

                case eExcelConditionalFormattingRuleType.Tomorrow:
                    return new ExcelConditionalFormattingTomorrow(address, ws, xr);

                case eExcelConditionalFormattingRuleType.Last7Days:
                    return new ExcelConditionalFormattingLast7Days(address, ws, xr);

                case eExcelConditionalFormattingRuleType.LastWeek:
                    return new ExcelConditionalFormattingLastWeek(address, ws, xr);

                case eExcelConditionalFormattingRuleType.ThisWeek:
                    return new ExcelConditionalFormattingThisWeek(address, ws, xr);

                case eExcelConditionalFormattingRuleType.NextWeek:
                    return new ExcelConditionalFormattingNextWeek(address, ws, xr);

                case eExcelConditionalFormattingRuleType.LastMonth:
                    return new ExcelConditionalFormattingLastMonth(address, ws, xr);

                case eExcelConditionalFormattingRuleType.ThisMonth:
                    return new ExcelConditionalFormattingThisMonth(address, ws, xr);

                case eExcelConditionalFormattingRuleType.NextMonth:
                    return new ExcelConditionalFormattingNextMonth(address, ws, xr);

                case eExcelConditionalFormattingRuleType.DuplicateValues:
                    return new ExcelConditionalFormattingDuplicateValues(address, ws, xr);

                case eExcelConditionalFormattingRuleType.Top:
                case eExcelConditionalFormattingRuleType.Bottom:
                case eExcelConditionalFormattingRuleType.TopPercent:
                case eExcelConditionalFormattingRuleType.BottomPercent:
                    return new ExcelConditionalFormattingTopBottomGroup(eType, address, ws, xr);

                case eExcelConditionalFormattingRuleType.AboveAverage:
                case eExcelConditionalFormattingRuleType.AboveOrEqualAverage:
                case eExcelConditionalFormattingRuleType.BelowOrEqualAverage:
                case eExcelConditionalFormattingRuleType.BelowAverage:
                    return new ExcelConditionalFormattingAverageGroup(eType, address, ws, xr);

                case eExcelConditionalFormattingRuleType.AboveStdDev:
                case eExcelConditionalFormattingRuleType.BelowStdDev:
                    return new ExcelConditionalFormattingStdDevGroup(eType, address, ws, xr);

                case eExcelConditionalFormattingRuleType.UniqueValues:
                    return new ExcelConditionalFormattingUniqueValues(address, ws, xr);

                case eExcelConditionalFormattingRuleType.DataBar:
                    return new ExcelConditionalFormattingDataBar(address, ws, xr);
            }

            throw new InvalidOperationException(
             string.Format(
             ExcelConditionalFormattingConstants.Errors.NonSupportedRuleType,
             eType.ToString()));
        }
    }
}