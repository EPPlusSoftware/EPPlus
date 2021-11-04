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
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;
using OfficeOpenXml.Utils;
using OfficeOpenXml.ConditionalFormatting.Contracts;

namespace OfficeOpenXml.ConditionalFormatting
{
	/// <summary>
	/// Factory class for ExcelConditionalFormatting
	/// </summary>
	internal static class ExcelConditionalFormattingRuleFactory
	{
		public static ExcelConditionalFormattingRule Create(
			eExcelConditionalFormattingRuleType type,
      ExcelAddress address,
      int priority,
			ExcelWorksheet worksheet,
			XmlNode itemElementNode)
		{
			Require.Argument(type);
      Require.Argument(address).IsNotNull("address");
      Require.Argument(worksheet).IsNotNull("worksheet");
			
			// According the conditional formatting rule type
			switch (type)
			{
        case eExcelConditionalFormattingRuleType.AboveAverage:
          return new ExcelConditionalFormattingAboveAverage(
            address,
            priority,
            worksheet,
            itemElementNode);

        case eExcelConditionalFormattingRuleType.AboveOrEqualAverage:
          return new ExcelConditionalFormattingAboveOrEqualAverage(
            address,
            priority,
            worksheet,
            itemElementNode);

        case eExcelConditionalFormattingRuleType.BelowAverage:
          return new ExcelConditionalFormattingBelowAverage(
            address,
            priority,
            worksheet,
            itemElementNode);

        case eExcelConditionalFormattingRuleType.BelowOrEqualAverage:
          return new ExcelConditionalFormattingBelowOrEqualAverage(
            address,
            priority,
            worksheet,
            itemElementNode);

        case eExcelConditionalFormattingRuleType.AboveStdDev:
          return new ExcelConditionalFormattingAboveStdDev(
            address,
            priority,
            worksheet,
            itemElementNode);

        case eExcelConditionalFormattingRuleType.BelowStdDev:
          return new ExcelConditionalFormattingBelowStdDev(
            address,
            priority,
            worksheet,
            itemElementNode);

        case eExcelConditionalFormattingRuleType.Bottom:
          return new ExcelConditionalFormattingBottom(
            address,
            priority,
            worksheet,
            itemElementNode);

        case eExcelConditionalFormattingRuleType.BottomPercent:
          return new ExcelConditionalFormattingBottomPercent(
            address,
            priority,
            worksheet,
            itemElementNode);

        case eExcelConditionalFormattingRuleType.Top:
          return new ExcelConditionalFormattingTop(
            address,
            priority,
            worksheet,
            itemElementNode);

        case eExcelConditionalFormattingRuleType.TopPercent:
          return new ExcelConditionalFormattingTopPercent(
            address,
            priority,
            worksheet,
            itemElementNode);

        case eExcelConditionalFormattingRuleType.Last7Days:
          return new ExcelConditionalFormattingLast7Days(
            address,
            priority,
            worksheet,
            itemElementNode);


        case eExcelConditionalFormattingRuleType.LastMonth:
          return new ExcelConditionalFormattingLastMonth(
            address,
            priority,
            worksheet,
            itemElementNode);

        case eExcelConditionalFormattingRuleType.LastWeek:
          return new ExcelConditionalFormattingLastWeek(
            address,
            priority,
            worksheet,
            itemElementNode);

        case eExcelConditionalFormattingRuleType.NextMonth:
          return new ExcelConditionalFormattingNextMonth(
            address,
            priority,
            worksheet,
            itemElementNode);

        case eExcelConditionalFormattingRuleType.NextWeek:
          return new ExcelConditionalFormattingNextWeek(
            address,
            priority,
            worksheet,
            itemElementNode);

        case eExcelConditionalFormattingRuleType.ThisMonth:
          return new ExcelConditionalFormattingThisMonth(
            address,
            priority,
            worksheet,
            itemElementNode);

        case eExcelConditionalFormattingRuleType.ThisWeek:
          return new ExcelConditionalFormattingThisWeek(
            address,
            priority,
            worksheet,
            itemElementNode);

        case eExcelConditionalFormattingRuleType.Today:
          return new ExcelConditionalFormattingToday(
            address,
            priority,
            worksheet,
            itemElementNode);

        case eExcelConditionalFormattingRuleType.Tomorrow:
          return new ExcelConditionalFormattingTomorrow(
            address,
            priority,
            worksheet,
            itemElementNode);

        case eExcelConditionalFormattingRuleType.Yesterday:
          return new ExcelConditionalFormattingYesterday(
            address,
            priority,
            worksheet,
            itemElementNode);

        case eExcelConditionalFormattingRuleType.BeginsWith:
          return new ExcelConditionalFormattingBeginsWith(
            address,
            priority,
            worksheet,
            itemElementNode);

        case eExcelConditionalFormattingRuleType.Between:
          return new ExcelConditionalFormattingBetween(
            address,
            priority,
            worksheet,
            itemElementNode);

        case eExcelConditionalFormattingRuleType.ContainsBlanks:
          return new ExcelConditionalFormattingContainsBlanks(
            address,
            priority,
            worksheet,
            itemElementNode);

        case eExcelConditionalFormattingRuleType.ContainsErrors:
          return new ExcelConditionalFormattingContainsErrors(
            address,
            priority,
            worksheet,
            itemElementNode);

        case eExcelConditionalFormattingRuleType.ContainsText:
          return new ExcelConditionalFormattingContainsText(
            address,
            priority,
            worksheet,
            itemElementNode);

        case eExcelConditionalFormattingRuleType.DuplicateValues:
          return new ExcelConditionalFormattingDuplicateValues(
            address,
            priority,
            worksheet,
            itemElementNode);

        case eExcelConditionalFormattingRuleType.EndsWith:
          return new ExcelConditionalFormattingEndsWith(
            address,
            priority,
            worksheet,
            itemElementNode);

        case eExcelConditionalFormattingRuleType.Equal:
          return new ExcelConditionalFormattingEqual(
            address,
            priority,
            worksheet,
            itemElementNode);

        case eExcelConditionalFormattingRuleType.Expression:
          return new ExcelConditionalFormattingExpression(
            address,
            priority,
            worksheet,
            itemElementNode);

        case eExcelConditionalFormattingRuleType.GreaterThan:
          return new ExcelConditionalFormattingGreaterThan(
            address,
            priority,
            worksheet,
            itemElementNode);

        case eExcelConditionalFormattingRuleType.GreaterThanOrEqual:
          return new ExcelConditionalFormattingGreaterThanOrEqual(
            address,
            priority,
            worksheet,
            itemElementNode);

        case eExcelConditionalFormattingRuleType.LessThan:
          return new ExcelConditionalFormattingLessThan(
            address,
            priority,
            worksheet,
            itemElementNode);

        case eExcelConditionalFormattingRuleType.LessThanOrEqual:
          return new ExcelConditionalFormattingLessThanOrEqual(
            address,
            priority,
            worksheet,
            itemElementNode);

        case eExcelConditionalFormattingRuleType.NotBetween:
          return new ExcelConditionalFormattingNotBetween(
            address,
            priority,
            worksheet,
            itemElementNode);

        case eExcelConditionalFormattingRuleType.NotContainsBlanks:
          return new ExcelConditionalFormattingNotContainsBlanks(
            address,
            priority,
            worksheet,
            itemElementNode);

        case eExcelConditionalFormattingRuleType.NotContainsErrors:
          return new ExcelConditionalFormattingNotContainsErrors(
            address,
            priority,
            worksheet,
            itemElementNode);

        case eExcelConditionalFormattingRuleType.NotContainsText:
          return new ExcelConditionalFormattingNotContainsText(
            address,
            priority,
            worksheet,
            itemElementNode);

        case eExcelConditionalFormattingRuleType.NotEqual:
          return new ExcelConditionalFormattingNotEqual(
            address,
            priority,
            worksheet,
            itemElementNode);

        case eExcelConditionalFormattingRuleType.UniqueValues:
          return new ExcelConditionalFormattingUniqueValues(
            address,
            priority,
            worksheet,
            itemElementNode);

        case eExcelConditionalFormattingRuleType.ThreeColorScale:
          return new ExcelConditionalFormattingThreeColorScale(
            address,
            priority,
            worksheet,
            itemElementNode);

        case eExcelConditionalFormattingRuleType.TwoColorScale:
          return new ExcelConditionalFormattingTwoColorScale(
            address,
            priority,
						worksheet,
						itemElementNode);
        case eExcelConditionalFormattingRuleType.ThreeIconSet:
          return new ExcelConditionalFormattingThreeIconSet(
            address,
            priority,
            worksheet,
            itemElementNode,
            null);
        case eExcelConditionalFormattingRuleType.FourIconSet:
          return new ExcelConditionalFormattingFourIconSet(
            address,
            priority,
            worksheet,
            itemElementNode,
            null);
        case eExcelConditionalFormattingRuleType.FiveIconSet:
          return new ExcelConditionalFormattingFiveIconSet(
            address,
            priority,
            worksheet,
            itemElementNode,
            null);
        case eExcelConditionalFormattingRuleType.DataBar:
          return new ExcelConditionalFormattingDataBar(
            eExcelConditionalFormattingRuleType.DataBar,
            address,
            priority,
            worksheet,
            itemElementNode,
            null);


        //TODO: Add DataBar
			}

			throw new InvalidOperationException(
        string.Format(
          ExcelConditionalFormattingConstants.Errors.NonSupportedRuleType,
          type.ToString()));
		}
	}
}