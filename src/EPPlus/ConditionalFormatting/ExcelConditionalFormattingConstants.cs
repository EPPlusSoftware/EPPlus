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
using System.Drawing;


namespace OfficeOpenXml.ConditionalFormatting
{
  /// <summary>
  /// The conditional formatting constants
  /// </summary>
  internal static class ExcelConditionalFormattingConstants
  {
    #region Errors
    internal class Errors
    {
      internal const string CommaSeparatedAddresses = @"Multiple addresses may not be commaseparated, use space instead";
      internal const string InvalidCfruleObject = @"The supplied item must inherit OfficeOpenXml.ConditionalFormatting.ExcelConditionalFormattingRule";
      internal const string InvalidConditionalFormattingObject = @"The supplied item must inherit OfficeOpenXml.ConditionalFormatting.ExcelConditionalFormatting";
      internal const string InvalidPriority = @"Invalid priority number. Must be bigger than zero";
      internal const string InvalidRemoveRuleOperation = @"Invalid remove rule operation";
      internal const string MissingCfvoNode = @"Missing 'cfvo' node in Conditional Formatting";
      internal const string MissingCfvoParentNode = @"Missing 'cfvo' parent node in Conditional Formatting";
      internal const string MissingConditionalFormattingNode = @"Missing 'conditionalFormatting' node in Conditional Formatting";
      internal const string MissingItemRuleList = @"Missing item with address '{0}' in Conditional Formatting Rule List";
      internal const string MissingPriorityAttribute = @"Missing 'priority' attribute in Conditional Formatting Rule";
      internal const string MissingRuleType = @"Missing eExcelConditionalFormattingRuleType Type in Conditional Formatting";
      internal const string MissingSqrefAttribute = @"Missing 'sqref' attribute in Conditional Formatting";
      internal const string MissingTypeAttribute = @"Missing 'type' attribute in Conditional Formatting Rule";
      internal const string MissingWorksheetNode = @"Missing 'worksheet' node";
      internal const string NonSupportedRuleType = @"Non supported conditionalFormattingType: {0}";
      internal const string UnexistentCfvoTypeAttribute = @"Unexistent eExcelConditionalFormattingValueObjectType attribute in Conditional Formatting";
      internal const string UnexistentOperatorTypeAttribute = @"Unexistent eExcelConditionalFormattingOperatorType attribute in Conditional Formatting";
      internal const string UnexistentTimePeriodTypeAttribute = @"Unexistent eExcelConditionalFormattingTimePeriodType attribute in Conditional Formatting";
      internal const string UnexpectedRuleTypeAttribute = @"Unexpected eExcelConditionalFormattingRuleType attribute in Conditional Formatting Rule";
      internal const string UnexpectedRuleTypeName = @"Unexpected eExcelConditionalFormattingRuleType TypeName in Conditional Formatting Rule";
      internal const string WrongNumberCfvoColorNodes = @"Wrong number of 'cfvo'/'color' nodes in Conditional Formatting Rule";
    }
    #endregion Errors

    #region Rule Type ST_CfType §18.18.12 (with small EPPlus changes)
    internal class RuleType
    {
      internal const string AboveAverage = "aboveAverage";
      internal const string BeginsWith = "beginsWith";
      internal const string CellIs = "cellIs";
      internal const string ColorScale = "colorScale";
      internal const string ContainsBlanks = "containsBlanks";
      internal const string ContainsErrors = "containsErrors";
      internal const string ContainsText = "containsText";
      internal const string DataBar = "dataBar";
      internal const string DuplicateValues = "duplicateValues";
      internal const string EndsWith = "endsWith";
      internal const string Expression = "expression";
      internal const string IconSet = "iconSet";
      internal const string NotContainsBlanks = "notContainsBlanks";
      internal const string NotContainsErrors = "notContainsErrors";
      internal const string NotContainsText = "notContainsText";
      internal const string TimePeriod = "timePeriod";
      internal const string Top10 = "top10";
      internal const string UniqueValues = "uniqueValues";

      // EPPlus Extended Types
      internal const string AboveOrEqualAverage = "aboveOrEqualAverage";
      internal const string AboveStdDev = "aboveStdDev";
      internal const string BelowAverage = "belowAverage";
      internal const string BelowOrEqualAverage = "belowOrEqualAverage";
      internal const string BelowStdDev = "belowStdDev";
      internal const string Between = "between";
      internal const string Bottom = "bottom";
      internal const string BottomPercent = "bottomPercent";
      internal const string Equal = "equal";
      internal const string GreaterThan = "greaterThan";
      internal const string GreaterThanOrEqual = "greaterThanOrEqual";
      internal const string IconSet3 = "iconSet3";
      internal const string IconSet4 = "iconSet4";
      internal const string IconSet5 = "iconSet5";
      internal const string Last7Days = "last7Days";
      internal const string LastMonth = "lastMonth";
      internal const string LastWeek = "lastWeek";
      internal const string LessThan = "lessThan";
      internal const string LessThanOrEqual = "lessThanOrEqual";
      internal const string NextMonth = "nextMonth";
      internal const string NextWeek = "nextWeek";
      internal const string NotBetween = "notBetween";
      internal const string NotEqual = "notEqual";
      internal const string ThisMonth = "thisMonth";
      internal const string ThisWeek = "thisWeek";
      internal const string ThreeColorScale = "threeColorScale";
      internal const string Today = "today";
      internal const string Tomorrow = "tomorrow";
      internal const string Top = "top";
      internal const string TopPercent = "topPercent";
      internal const string TwoColorScale = "twoColorScale";
      internal const string Yesterday = "yesterday";
    }
    #endregion Rule Type ST_CfType §18.18.12 (with small EPPlus changes)

    #region Colors
    internal class Colors
    {
      internal static readonly Color CfvoLowValue = Color.FromArgb(0xFF,0xF8,0x69,0x6B);
      internal static readonly Color CfvoMiddleValue = Color.FromArgb(0xFF,0xFF,0xEB,0x84);
      internal static readonly Color CfvoHighValue = Color.FromArgb(0xFF,0x63,0xBE,0x7B);
    }

    #endregion Colors
    }
}