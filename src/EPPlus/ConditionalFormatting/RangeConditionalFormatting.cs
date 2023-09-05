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
using OfficeOpenXml.Utils;
using OfficeOpenXml.ConditionalFormatting.Contracts;
using System.Collections.Generic;
using OfficeOpenXml.Core.RangeQuadTree;
using System;

namespace OfficeOpenXml.ConditionalFormatting
{
    internal class RangeConditionalFormatting
      : IRangeConditionalFormatting
    {
        #region Public Properties
        public ExcelWorksheet _worksheet;
        public ExcelAddress _address;
        #endregion Public Properties

        #region Constructors
        public RangeConditionalFormatting(
          ExcelWorksheet worksheet,
          ExcelAddress address)
        {
            Require.Argument(worksheet).IsNotNull("worksheet");
            Require.Argument(address).IsNotNull("address");

            _worksheet = worksheet;
            _address = address;
        }
        #endregion Constructors

        #region Conditional Formatting Rule Types
        /// <summary>
        /// Add AboveOrEqualAverage Conditional Formatting
        /// </summary>
        ///  <returns></returns>
        public IExcelConditionalFormattingAverageGroup AddAboveAverage()
        {
            return _worksheet.ConditionalFormatting.AddAboveAverage(
              _address);
        }

        /// <summary>
        /// Add AboveOrEqualAverage Conditional Formatting
        /// </summary>
        ///  <returns></returns>
        public IExcelConditionalFormattingAverageGroup AddAboveOrEqualAverage()
        {
            return _worksheet.ConditionalFormatting.AddAboveOrEqualAverage(
              _address);
        }

        /// <summary>
        /// Add BelowOrEqualAverage Conditional Formatting
        /// </summary>
        ///  <returns></returns>
        public IExcelConditionalFormattingAverageGroup AddBelowAverage()
        {
            return _worksheet.ConditionalFormatting.AddBelowAverage(
              _address);
        }

        /// <summary>
        /// Add BelowOrEqualAverage Conditional Formatting
        /// </summary>
        ///  <returns></returns>
        public IExcelConditionalFormattingAverageGroup AddBelowOrEqualAverage()
        {
            return _worksheet.ConditionalFormatting.AddBelowOrEqualAverage(
              _address);
        }

        /// <summary>
        /// Add AboveStdDev Conditional Formatting
        /// </summary>
        ///  <returns></returns>
        public IExcelConditionalFormattingStdDevGroup AddAboveStdDev()
        {
            return _worksheet.ConditionalFormatting.AddAboveStdDev(
              _address);
        }

        /// <summary>
        /// Add BelowStdDev Conditional Formatting
        /// </summary>
        ///  <returns></returns>
        public IExcelConditionalFormattingStdDevGroup AddBelowStdDev()
        {
            return _worksheet.ConditionalFormatting.AddBelowStdDev(
              _address);
        }

        /// <summary>
        /// Add Bottom Conditional Formatting
        /// </summary>
        ///  <returns></returns>
        public IExcelConditionalFormattingTopBottomGroup AddBottom()
        {
            return _worksheet.ConditionalFormatting.AddBottom(
              _address);
        }

        /// <summary>
        /// Add BottomPercent Conditional Formatting
        /// </summary>
        ///  <returns></returns>
        public IExcelConditionalFormattingTopBottomGroup AddBottomPercent()
        {
            return _worksheet.ConditionalFormatting.AddBottomPercent(
              _address);
        }

        /// <summary>
        /// Add Top Conditional Formatting
        /// </summary>
        ///  <returns></returns>
        public IExcelConditionalFormattingTopBottomGroup AddTop()
        {
            return _worksheet.ConditionalFormatting.AddTop(
              _address);
        }

        /// <summary>
        /// Add TopPercent Conditional Formatting
        /// </summary>
        ///  <returns></returns>
        public IExcelConditionalFormattingTopBottomGroup AddTopPercent()
        {
            return _worksheet.ConditionalFormatting.AddTopPercent(
              _address);
        }

        /// <summary>
        /// Add Last7Days Conditional Formatting
        /// </summary>
        ///  <returns></returns>
        public IExcelConditionalFormattingTimePeriodGroup AddLast7Days()
        {
            return _worksheet.ConditionalFormatting.AddLast7Days(
              _address);
        }

        /// <summary>
        /// Add LastMonth Conditional Formatting
        /// </summary>
        ///  <returns></returns>
        public IExcelConditionalFormattingTimePeriodGroup AddLastMonth()
        {
            return _worksheet.ConditionalFormatting.AddLastMonth(
              _address);
        }

        /// <summary>
        /// Add LastWeek Conditional Formatting
        /// </summary>
        ///  <returns></returns>
        public IExcelConditionalFormattingTimePeriodGroup AddLastWeek()
        {
            return _worksheet.ConditionalFormatting.AddLastWeek(
              _address);
        }

        /// <summary>
        /// Add NextMonth Conditional Formatting
        /// </summary>
        ///  <returns></returns>
        public IExcelConditionalFormattingTimePeriodGroup AddNextMonth()
        {
            return _worksheet.ConditionalFormatting.AddNextMonth(
              _address);
        }

        /// <summary>
        /// Add NextWeek Conditional Formatting
        /// </summary>
        ///  <returns></returns>
        public IExcelConditionalFormattingTimePeriodGroup AddNextWeek()
        {
            return _worksheet.ConditionalFormatting.AddNextWeek(
              _address);
        }

        /// <summary>
        /// Add ThisMonth Conditional Formatting
        /// </summary>
        ///  <returns></returns>
        public IExcelConditionalFormattingTimePeriodGroup AddThisMonth()
        {
            return _worksheet.ConditionalFormatting.AddThisMonth(
              _address);
        }

        /// <summary>
        /// Add ThisWeek Conditional Formatting
        /// </summary>
        ///  <returns></returns>
        public IExcelConditionalFormattingTimePeriodGroup AddThisWeek()
        {
            return _worksheet.ConditionalFormatting.AddThisWeek(
              _address);
        }

        /// <summary>
        /// Add Today Conditional Formatting
        /// </summary>
        ///  <returns></returns>
        public IExcelConditionalFormattingTimePeriodGroup AddToday()
        {
            return _worksheet.ConditionalFormatting.AddToday(
              _address);
        }

        /// <summary>
        /// Add Tomorrow Conditional Formatting
        /// </summary>
        ///  <returns></returns>
        public IExcelConditionalFormattingTimePeriodGroup AddTomorrow()
        {
            return _worksheet.ConditionalFormatting.AddTomorrow(
              _address);
        }

        /// <summary>
        /// Add Yesterday Conditional Formatting
        /// </summary>
        ///  <returns></returns>
        public IExcelConditionalFormattingTimePeriodGroup AddYesterday()
        {
            return _worksheet.ConditionalFormatting.AddYesterday(
              _address);
        }

        /// <summary>
        /// Add BeginsWith Conditional Formatting
        /// </summary>
        ///  <returns></returns>
        public IExcelConditionalFormattingBeginsWith AddBeginsWith()
        {
            return _worksheet.ConditionalFormatting.AddBeginsWith(
              _address);
        }

        /// <summary>
        /// Add Between Conditional Formatting
        /// </summary>
        ///  <returns></returns>
        public IExcelConditionalFormattingBetween AddBetween()
        {
            return _worksheet.ConditionalFormatting.AddBetween(
              _address);
        }

        /// <summary>
        /// Add ContainsBlanks Conditional Formatting
        /// </summary>
        ///  <returns></returns>
        public IExcelConditionalFormattingContainsBlanks AddContainsBlanks()
        {
            return _worksheet.ConditionalFormatting.AddContainsBlanks(
              _address);
        }

        /// <summary>
        /// Add ContainsErrors Conditional Formatting
        /// </summary>
        ///  <returns></returns>
        public IExcelConditionalFormattingContainsErrors AddContainsErrors()
        {
            return _worksheet.ConditionalFormatting.AddContainsErrors(
              _address);
        }

        /// <summary>
        /// Add ContainsText Conditional Formatting
        /// </summary>
        ///  <returns></returns>
        public IExcelConditionalFormattingContainsText AddContainsText()
        {
            return _worksheet.ConditionalFormatting.AddContainsText(
              _address);
        }

        /// <summary>
        /// Add DuplicateValues Conditional Formatting
        /// </summary>
        ///  <returns></returns>
        public IExcelConditionalFormattingDuplicateValues AddDuplicateValues()
        {
            return _worksheet.ConditionalFormatting.AddDuplicateValues(
              _address);
        }

        /// <summary>
        /// Add EndsWith Conditional Formatting
        /// </summary>
        ///  <returns></returns>
        public IExcelConditionalFormattingEndsWith AddEndsWith()
        {
            return _worksheet.ConditionalFormatting.AddEndsWith(
              _address);
        }

        /// <summary>
        /// Add Equal Conditional Formatting
        /// </summary>
        ///  <returns></returns>
        public IExcelConditionalFormattingEqual AddEqual()
        {
            return _worksheet.ConditionalFormatting.AddEqual(
              _address);
        }

        /// <summary>
        /// Add Expression Conditional Formatting
        /// </summary>
        ///  <returns></returns>
        public IExcelConditionalFormattingExpression AddExpression()
        {
            return _worksheet.ConditionalFormatting.AddExpression(
              _address);
        }

        /// <summary>
        /// Add GreaterThan Conditional Formatting
        /// </summary>
        ///  <returns></returns>
        public IExcelConditionalFormattingGreaterThan AddGreaterThan()
        {
            return _worksheet.ConditionalFormatting.AddGreaterThan(
              _address);
        }

        /// <summary>
        /// Add GreaterThanOrEqual Conditional Formatting
        /// </summary>
        ///  <returns></returns>
        public IExcelConditionalFormattingGreaterThanOrEqual AddGreaterThanOrEqual()
        {
            return _worksheet.ConditionalFormatting.AddGreaterThanOrEqual(
              _address);
        }

        /// <summary>
        /// Add LessThan Conditional Formatting
        /// </summary>
        ///  <returns></returns>
        public IExcelConditionalFormattingLessThan AddLessThan()
        {
            return _worksheet.ConditionalFormatting.AddLessThan(
              _address);
        }

        /// <summary>
        /// Add LessThanOrEqual Conditional Formatting
        /// </summary>
        ///  <returns></returns>
        public IExcelConditionalFormattingLessThanOrEqual AddLessThanOrEqual()
        {
            return _worksheet.ConditionalFormatting.AddLessThanOrEqual(
              _address);
        }

        /// <summary>
        /// Add NotBetween Conditional Formatting
        /// </summary>
        ///  <returns></returns>
        public IExcelConditionalFormattingNotBetween AddNotBetween()
        {
            return _worksheet.ConditionalFormatting.AddNotBetween(
              _address);
        }

        /// <summary>
        /// Add NotContainsBlanks Conditional Formatting
        /// </summary>
        ///  <returns></returns>
        public IExcelConditionalFormattingNotContainsBlanks AddNotContainsBlanks()
        {
            return _worksheet.ConditionalFormatting.AddNotContainsBlanks(
              _address);
        }

        /// <summary>
        /// Add NotContainsErrors Conditional Formatting
        /// </summary>
        ///  <returns></returns>
        public IExcelConditionalFormattingNotContainsErrors AddNotContainsErrors()
        {
            return _worksheet.ConditionalFormatting.AddNotContainsErrors(
              _address);
        }

        /// <summary>
        /// Add NotContainsText Conditional Formatting
        /// </summary>
        ///  <returns></returns>
        public IExcelConditionalFormattingNotContainsText AddNotContainsText()
        {
            return _worksheet.ConditionalFormatting.AddNotContainsText(
              _address);
        }

        /// <summary>
        /// Add NotEqual Conditional Formatting
        /// </summary>
        ///  <returns></returns>
        public IExcelConditionalFormattingNotEqual AddNotEqual()
        {
            return _worksheet.ConditionalFormatting.AddNotEqual(
              _address);
        }

        /// <summary>
        /// Add UniqueValues Conditional Formatting
        /// </summary>
        ///  <returns></returns>
        public IExcelConditionalFormattingUniqueValues AddUniqueValues()
        {
            return _worksheet.ConditionalFormatting.AddUniqueValues(
              _address);
        }

        /// <summary>
        /// Add ThreeColorScale Conditional Formatting
        /// </summary>
        ///  <returns></returns>
        public IExcelConditionalFormattingThreeColorScale AddThreeColorScale()
        {
            return (IExcelConditionalFormattingThreeColorScale)(_worksheet.ConditionalFormatting.AddRule(
              eExcelConditionalFormattingRuleType.ThreeColorScale,
              _address));
        }

        /// <summary>
        /// Add TwoColorScale Conditional Formatting
        /// </summary>
        ///  <returns></returns>
        public IExcelConditionalFormattingTwoColorScale AddTwoColorScale()
        {
            return (IExcelConditionalFormattingTwoColorScale)(_worksheet.ConditionalFormatting.AddRule(
              eExcelConditionalFormattingRuleType.TwoColorScale,
              _address));
        }

        /// <summary>
        /// Adds a ThreeIconSet rule 
        /// </summary>
        /// <param name="IconSet"></param>
        /// <returns></returns>
        public IExcelConditionalFormattingThreeIconSet<eExcelconditionalFormatting3IconsSetType> AddThreeIconSet(eExcelconditionalFormatting3IconsSetType IconSet)
        {
            var rule = (IExcelConditionalFormattingThreeIconSet<eExcelconditionalFormatting3IconsSetType>)(_worksheet.ConditionalFormatting.AddRule(
              eExcelConditionalFormattingRuleType.ThreeIconSet,
              _address));
            rule.IconSet = IconSet;
            return rule;
        }

        /// <summary>
        /// Adds a FourIconSet rule 
        /// </summary>
        /// <param name="IconSet"></param>
        /// <returns></returns>
        public IExcelConditionalFormattingFourIconSet<eExcelconditionalFormatting4IconsSetType> AddFourIconSet(eExcelconditionalFormatting4IconsSetType IconSet)
        {
            var rule = (IExcelConditionalFormattingFourIconSet<eExcelconditionalFormatting4IconsSetType>)(_worksheet.ConditionalFormatting.AddRule(
              eExcelConditionalFormattingRuleType.FourIconSet,
              _address));
            rule.IconSet = IconSet;
            return rule;
        }

        /// <summary>
        /// Adds a FiveIconSet rule 
        /// </summary>
        /// <param name="IconSet"></param>
        /// <returns></returns>
        public IExcelConditionalFormattingFiveIconSet AddFiveIconSet(eExcelconditionalFormatting5IconsSetType IconSet)
        {
            var rule = (IExcelConditionalFormattingFiveIconSet)(_worksheet.ConditionalFormatting.AddRule(
              eExcelConditionalFormattingRuleType.FiveIconSet,
              _address));
            rule.IconSet = IconSet;
            return rule;
        }

        /// <summary>
        /// Adds a Databar rule 
        /// </summary>
        /// <param name="Color">The color of the databar</param>
        /// <returns></returns>
        public IExcelConditionalFormattingDataBarGroup AddDatabar(System.Drawing.Color Color)
        {
            var rule = (IExcelConditionalFormattingDataBarGroup)(_worksheet.ConditionalFormatting.AddRule(
              eExcelConditionalFormattingRuleType.DataBar,
              _address));
            rule.Color = Color;
            return rule;
        }
        #endregion Conditional Formatting Rule Types

        #region Public Methods
        public Dictionary<string, List<ExcelConditionalFormattingRule>> GetConditionalFormattings()
        {
            var conditionalFormattingRules = new QuadTree<ExcelConditionalFormattingRule>(_address);
            var retDikt = new Dictionary<string, List<ExcelConditionalFormattingRule>>();

            foreach (var cf in _worksheet.ConditionalFormatting)
            {
                var intersects = conditionalFormattingRules.GetIntersectRanges(new QuadRange(cf.Address));

                foreach (var intersect in intersects)
                {
                    var address = new ExcelAddress(intersect.FromRow, intersect.FromCol, intersect.ToRow, intersect.ToCol);

                    if (!retDikt.ContainsKey(address.Address))
                    {
                        retDikt.Add(address.Address, new List<ExcelConditionalFormattingRule>());
                    }

                    retDikt[address.Address].Add((ExcelConditionalFormattingRule)cf);
                }
            }

            return retDikt;
        }
        #endregion Public Methods
        #endregion Conditional Formatting Rule Types
    }
}