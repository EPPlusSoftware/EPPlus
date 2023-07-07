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
using System;

namespace OfficeOpenXml.ConditionalFormatting
{
    /// <summary>
    /// Functions related to the ExcelConditionalFormattingRule
    /// </summary>
    internal static class ExcelConditionalFormattingRuleType
    {
        /// <summary>
        /// 
        /// </summary>
        /// <param name="type"></param>
        /// <returns></returns>
        public static string GetAttributeByType(
          eExcelConditionalFormattingRuleType type)
        {
            switch (type)
            {
                case eExcelConditionalFormattingRuleType.AboveAverage:
                case eExcelConditionalFormattingRuleType.AboveOrEqualAverage:
                case eExcelConditionalFormattingRuleType.BelowAverage:
                case eExcelConditionalFormattingRuleType.BelowOrEqualAverage:
                case eExcelConditionalFormattingRuleType.AboveStdDev:
                case eExcelConditionalFormattingRuleType.BelowStdDev:
                    return ExcelConditionalFormattingConstants.RuleType.AboveAverage;

                case eExcelConditionalFormattingRuleType.Bottom:
                case eExcelConditionalFormattingRuleType.BottomPercent:
                case eExcelConditionalFormattingRuleType.Top:
                case eExcelConditionalFormattingRuleType.TopPercent:
                    return ExcelConditionalFormattingConstants.RuleType.Top10;

                case eExcelConditionalFormattingRuleType.Last7Days:
                case eExcelConditionalFormattingRuleType.LastMonth:
                case eExcelConditionalFormattingRuleType.LastWeek:
                case eExcelConditionalFormattingRuleType.NextMonth:
                case eExcelConditionalFormattingRuleType.NextWeek:
                case eExcelConditionalFormattingRuleType.ThisMonth:
                case eExcelConditionalFormattingRuleType.ThisWeek:
                case eExcelConditionalFormattingRuleType.Today:
                case eExcelConditionalFormattingRuleType.Tomorrow:
                case eExcelConditionalFormattingRuleType.Yesterday:
                    return ExcelConditionalFormattingConstants.RuleType.TimePeriod;

                case eExcelConditionalFormattingRuleType.Between:
                case eExcelConditionalFormattingRuleType.Equal:
                case eExcelConditionalFormattingRuleType.GreaterThan:
                case eExcelConditionalFormattingRuleType.GreaterThanOrEqual:
                case eExcelConditionalFormattingRuleType.LessThan:
                case eExcelConditionalFormattingRuleType.LessThanOrEqual:
                case eExcelConditionalFormattingRuleType.NotBetween:
                case eExcelConditionalFormattingRuleType.NotEqual:
                    return ExcelConditionalFormattingConstants.RuleType.CellIs;

                case eExcelConditionalFormattingRuleType.ThreeIconSet:
                case eExcelConditionalFormattingRuleType.FourIconSet:
                case eExcelConditionalFormattingRuleType.FiveIconSet:
                    return ExcelConditionalFormattingConstants.RuleType.IconSet;

                case eExcelConditionalFormattingRuleType.ThreeColorScale:
                case eExcelConditionalFormattingRuleType.TwoColorScale:
                    return ExcelConditionalFormattingConstants.RuleType.ColorScale;

                case eExcelConditionalFormattingRuleType.BeginsWith:
                    return ExcelConditionalFormattingConstants.RuleType.BeginsWith;

                case eExcelConditionalFormattingRuleType.ContainsBlanks:
                    return ExcelConditionalFormattingConstants.RuleType.ContainsBlanks;

                case eExcelConditionalFormattingRuleType.ContainsErrors:
                    return ExcelConditionalFormattingConstants.RuleType.ContainsErrors;

                case eExcelConditionalFormattingRuleType.ContainsText:
                    return ExcelConditionalFormattingConstants.RuleType.ContainsText;

                case eExcelConditionalFormattingRuleType.DuplicateValues:
                    return ExcelConditionalFormattingConstants.RuleType.DuplicateValues;

                case eExcelConditionalFormattingRuleType.EndsWith:
                    return ExcelConditionalFormattingConstants.RuleType.EndsWith;

                case eExcelConditionalFormattingRuleType.Expression:
                    return ExcelConditionalFormattingConstants.RuleType.Expression;

                case eExcelConditionalFormattingRuleType.NotContainsBlanks:
                    return ExcelConditionalFormattingConstants.RuleType.NotContainsBlanks;

                case eExcelConditionalFormattingRuleType.NotContainsErrors:
                    return ExcelConditionalFormattingConstants.RuleType.NotContainsErrors;

                case eExcelConditionalFormattingRuleType.NotContainsText:
                    return ExcelConditionalFormattingConstants.RuleType.NotContainsText;

                case eExcelConditionalFormattingRuleType.UniqueValues:
                    return ExcelConditionalFormattingConstants.RuleType.UniqueValues;

                case eExcelConditionalFormattingRuleType.DataBar:
                    return ExcelConditionalFormattingConstants.RuleType.DataBar;
            }

            throw new Exception(
              ExcelConditionalFormattingConstants.Errors.MissingRuleType);
        }
    }
}