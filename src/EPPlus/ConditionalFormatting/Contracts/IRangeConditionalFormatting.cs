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
using OfficeOpenXml.ConditionalFormatting.Contracts;
using System.Collections.Generic;
using System.Drawing;

namespace OfficeOpenXml.ConditionalFormatting
{
    /// <summary>
    /// Provides functionality for adding Conditional Formatting to a range (<see cref="ExcelRangeBase"/>).
    /// Each method will return a configurable condtional formatting type.
    /// </summary>
    public interface IRangeConditionalFormatting
    {
        /// <summary>
        /// Adds an Above Average rule to the range
        /// </summary>
        /// <returns></returns>
        IExcelConditionalFormattingAverageGroup AddAboveAverage();

        /// <summary>
        /// Adds an Above Or Equal Average rule to the range
        /// </summary>
        /// <returns></returns>
        IExcelConditionalFormattingAverageGroup AddAboveOrEqualAverage();

        /// <summary>
        /// Adds a Below Average rule to the range
        /// </summary>
        /// <returns></returns>
        IExcelConditionalFormattingAverageGroup AddBelowAverage();

        /// <summary>
        /// Adds a Below Or Equal Average rule to the range
        /// </summary>
        /// <returns></returns>
        IExcelConditionalFormattingAverageGroup AddBelowOrEqualAverage();

        /// <summary>
        /// Adds an Above StdDev rule to the range
        /// </summary>
        /// <returns></returns>
        IExcelConditionalFormattingStdDevGroup AddAboveStdDev();

        /// <summary>
        /// Adds an Below StdDev rule to the range
        /// </summary>
        /// <returns></returns>
        IExcelConditionalFormattingStdDevGroup AddBelowStdDev();

        /// <summary>
        /// Adds a Bottom rule to the range
        /// </summary>
        /// <returns></returns>
        IExcelConditionalFormattingTopBottomGroup AddBottom();

        /// <summary>
        /// Adds a Bottom Percent rule to the range
        /// </summary>
        /// <returns></returns>
        IExcelConditionalFormattingTopBottomGroup AddBottomPercent();

        /// <summary>
        /// Adds a Top rule to the range
        /// </summary>
        /// <returns></returns>
        IExcelConditionalFormattingTopBottomGroup AddTop();

        /// <summary>
        /// Adds a Top Percent rule to the range
        /// </summary>
        /// <returns></returns>
        IExcelConditionalFormattingTopBottomGroup AddTopPercent();

        /// <summary>
        /// Adds a Last 7 Days rule to the range
        /// </summary>
        /// <returns></returns>
        IExcelConditionalFormattingTimePeriodGroup AddLast7Days();

        /// <summary>
        /// Adds a Last Month rule to the range
        /// </summary>
        /// <returns></returns>
        IExcelConditionalFormattingTimePeriodGroup AddLastMonth();

        /// <summary>
        /// Adds a Last Week rule to the range
        /// </summary>
        /// <returns></returns>
        IExcelConditionalFormattingTimePeriodGroup AddLastWeek();

        /// <summary>
        /// Adds a Next Month rule to the range
        /// </summary>
        /// <returns></returns>
        IExcelConditionalFormattingTimePeriodGroup AddNextMonth();

        /// <summary>
        /// Adds a Next Week rule to the range
        /// </summary>
        /// <returns></returns>
        IExcelConditionalFormattingTimePeriodGroup AddNextWeek();

        /// <summary>
        /// Adds a This Month rule to the range
        /// </summary>
        /// <returns></returns>
        IExcelConditionalFormattingTimePeriodGroup AddThisMonth();

        /// <summary>
        /// Adds a This Week rule to the range
        /// </summary>
        /// <returns></returns>
        IExcelConditionalFormattingTimePeriodGroup AddThisWeek();

        /// <summary>
        /// Adds a Today rule to the range
        /// </summary>
        /// <returns></returns>
        IExcelConditionalFormattingTimePeriodGroup AddToday();

        /// <summary>
        /// Adds a Tomorrow rule to the range
        /// </summary>
        /// <returns></returns>
        IExcelConditionalFormattingTimePeriodGroup AddTomorrow();

        /// <summary>
        /// Adds an Yesterday rule to the range
        /// </summary>
        /// <returns></returns>
        IExcelConditionalFormattingTimePeriodGroup AddYesterday();

        /// <summary>
        /// Adds a Begins With rule to the range
        /// </summary>
        /// <returns></returns>
        IExcelConditionalFormattingBeginsWith AddBeginsWith();

        /// <summary>
        /// Adds a Between rule to the range
        /// </summary>
        /// <returns></returns>
        IExcelConditionalFormattingBetween AddBetween();

        /// <summary>
        /// Adds a ContainsBlanks rule to the range
        /// </summary>
        /// <returns></returns>
        IExcelConditionalFormattingContainsBlanks AddContainsBlanks();

        /// <summary>
        /// Adds a ContainsErrors rule to the range
        /// </summary>
        /// <returns></returns>
        IExcelConditionalFormattingContainsErrors AddContainsErrors();

        /// <summary>
        /// Adds a ContainsText rule to the range
        /// </summary>
        /// <returns></returns>
        IExcelConditionalFormattingContainsText AddContainsText();

        /// <summary>
        /// Adds a DuplicateValues rule to the range
        /// </summary>
        /// <returns></returns>
        IExcelConditionalFormattingDuplicateValues AddDuplicateValues();

        /// <summary>
        /// Adds an EndsWith rule to the range
        /// </summary>
        /// <returns></returns>
        IExcelConditionalFormattingEndsWith AddEndsWith();

        /// <summary>
        /// Adds an Equal rule to the range
        /// </summary>
        /// <returns></returns>
        IExcelConditionalFormattingEqual AddEqual();

        /// <summary>
        /// Adds an Expression rule to the range
        /// </summary>
        /// <returns></returns>
        IExcelConditionalFormattingExpression AddExpression();

        /// <summary>
        /// Adds a GreaterThan rule to the range
        /// </summary>
        /// <returns></returns>
        IExcelConditionalFormattingGreaterThan AddGreaterThan();

        /// <summary>
        /// Adds a GreaterThanOrEqual rule to the range
        /// </summary>
        /// <returns></returns>
        IExcelConditionalFormattingGreaterThanOrEqual AddGreaterThanOrEqual();

        /// <summary>
        /// Adds a LessThan rule to the range
        /// </summary>
        /// <returns></returns>
        IExcelConditionalFormattingLessThan AddLessThan();

        /// <summary>
        /// Adds a LessThanOrEqual rule to the range
        /// </summary>
        /// <returns></returns>
        IExcelConditionalFormattingLessThanOrEqual AddLessThanOrEqual();

        /// <summary>
        /// Adds a NotBetween rule to the range
        /// </summary>
        /// <returns></returns>
        IExcelConditionalFormattingNotBetween AddNotBetween();

        /// <summary>
        /// Adds a NotContainsBlanks rule to the range
        /// </summary>
        /// <returns></returns>
        IExcelConditionalFormattingNotContainsBlanks AddNotContainsBlanks();

        /// <summary>
        /// Adds a NotContainsErrors rule to the range
        /// </summary>
        /// <returns></returns>
        IExcelConditionalFormattingNotContainsErrors AddNotContainsErrors();

        /// <summary>
        /// Adds a NotContainsText rule to the range
        /// </summary>
        /// <returns></returns>
        IExcelConditionalFormattingNotContainsText AddNotContainsText();

        /// <summary>
        /// Adds a NotEqual rule to the range
        /// </summary>
        /// <returns></returns>
        IExcelConditionalFormattingNotEqual AddNotEqual();

        /// <summary>
        /// Adds an UniqueValues rule to the range
        /// </summary>
        /// <returns></returns>
        IExcelConditionalFormattingUniqueValues AddUniqueValues();

        /// <summary>
        /// Adds a <see cref="ExcelConditionalFormattingThreeColorScale"/> to the range
        /// </summary>
        /// <returns></returns>
        IExcelConditionalFormattingThreeColorScale AddThreeColorScale();

        /// <summary>
        /// Adds a <see cref="ExcelConditionalFormattingTwoColorScale"/> to the range
        /// </summary>
        /// <returns></returns>
        IExcelConditionalFormattingTwoColorScale AddTwoColorScale();

        /// <summary>
        /// Adds a <see cref="IExcelConditionalFormattingThreeIconSet&lt;eExcelconditionalFormatting3IconsSetType&gt;"/> to the range
        /// </summary>
        /// <param name="IconSet"></param>
        /// <returns></returns>
        IExcelConditionalFormattingThreeIconSet<eExcelconditionalFormatting3IconsSetType> AddThreeIconSet(eExcelconditionalFormatting3IconsSetType IconSet);
        /// <summary>
        /// Adds a <see cref="IExcelConditionalFormattingFourIconSet&lt;eExcelconditionalFormatting4IconsSetType&gt;"/> to the range
        /// </summary>
        /// <param name="IconSet"></param>
        /// <returns></returns>
        IExcelConditionalFormattingFourIconSet<eExcelconditionalFormatting4IconsSetType> AddFourIconSet(eExcelconditionalFormatting4IconsSetType IconSet);
        /// <summary>
        /// Adds a <see cref="IExcelConditionalFormattingFiveIconSet"/> to the range
        /// </summary>
        /// <param name="IconSet"></param>
        /// <returns></returns>
        IExcelConditionalFormattingFiveIconSet AddFiveIconSet(eExcelconditionalFormatting5IconsSetType IconSet);
        /// <summary>
        /// Adds a <see cref="IExcelConditionalFormattingDataBarGroup"/> to the range
        /// </summary>
        /// <param name="color"></param>
        /// <returns></returns>
        IExcelConditionalFormattingDataBarGroup AddDatabar(Color color);

        /// <summary>
        /// Get a list of all conditional formatting rules that exist on cells in the range
        /// </summary>
        /// <returns></returns>
        List<ExcelConditionalFormattingRule> GetConditionalFormattings();
    }
}