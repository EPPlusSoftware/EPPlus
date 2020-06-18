using OfficeOpenXml.ConditionalFormatting.Contracts;
using System;
using System.Collections.Generic;
using System.Text;

namespace OfficeOpenXml.ConditionalFormatting
{
    /// <summary>
    /// Provides a simple way to type cast a conditional formatting object top its top level class.
    /// </summary>
    public class ExcelConditionalFormattingAsType
    {
        IExcelConditionalFormattingRule _rule;
        internal ExcelConditionalFormattingAsType(IExcelConditionalFormattingRule rule)
        {
            _rule = rule;
        }

        /// <summary>
        /// Converts the conditional formatting object to it's top level or another nested class.        
        /// </summary>
        /// <typeparam name="T">The type of conditional formatting object. T must be inherited from IExcelConditionalFormattingRule</typeparam>
        /// <returns>The conditional formatting rule as type T</returns>
        public T Type<T>() where T : IExcelConditionalFormattingRule
        {
            if(_rule is T t)
            {
                return t;
            }
            return default;
        }
        /// <summary>
        /// Returns the conditional formatting object as an Average rule
        /// If this object is not of type AboveAverage, AboveOrEqualAverage, BelowAverage or BelowOrEqualAverage, null will be returned
        /// </summary>
        /// <returns>The conditional formatting rule as an Average rule</returns>
        public IExcelConditionalFormattingAverageGroup Average
        {
            get
            {
                return _rule as IExcelConditionalFormattingAverageGroup;
            }
        }
        /// <summary>
        /// Returns the conditional formatting object as a StdDev rule
        /// If this object is not of type AboveStdDev or BelowStdDev, null will be returned
        /// </summary>
        /// <returns>The conditional formatting object as a StdDev rule</returns>
        public IExcelConditionalFormattingStdDevGroup StdDev
        {
            get
            {
                return _rule as IExcelConditionalFormattingStdDevGroup;
            }
        }
        /// <summary>
        /// Returns the conditional formatting object as a TopBottom rule
        /// If this object is not of type Bottom, BottomPercent, Top or TopPercent, null will be returned
        /// </summary>
        /// <returns>The conditional formatting object as a TopBottom rule</returns>
        public IExcelConditionalFormattingTopBottomGroup TopBottom
        {
            get
            {
                return _rule as IExcelConditionalFormattingTopBottomGroup;
            }
        }
        /// <summary>
        /// Returns the conditional formatting object as a DateTimePeriod rule
        /// If this object is not of type Last7Days, LastMonth, LastWeek, NextMonth, NextWeek, ThisMonth, ThisWeek, Today, Tomorrow or Yesterday, null will be returned
        /// </summary>
        /// <returns>The conditional formatting object as a DateTimePeriod rule</returns>
        public IExcelConditionalFormattingTimePeriodGroup DateTimePeriod
        {
            get
            {
                return _rule as IExcelConditionalFormattingTimePeriodGroup;
            }
        }
        /// <summary>
        /// Returns the conditional formatting object as a Between rule
        /// If this object is not of type Between, null will be returned
        /// </summary>
        /// <returns>The conditional formatting object as a Between rule</returns>
        public IExcelConditionalFormattingBetween Between
        {
            get
            {
                return _rule as IExcelConditionalFormattingBetween;
            }
        }
        /// <summary>
        /// Returns the conditional formatting object as a ContainsBlanks rule
        /// If this object is not of type ContainsBlanks, null will be returned
        /// </summary>
        /// <returns>The conditional formatting object as a ContainsBlanks rule</returns>
        public IExcelConditionalFormattingContainsBlanks ContainsBlanks
        {
            get
            {
                return _rule as IExcelConditionalFormattingContainsBlanks;
            }
        }
        /// <summary>
        /// Returns the conditional formatting object as a ContainsErrors rule
        /// If this object is not of type ContainsErrors, null will be returned
        /// </summary>
        /// <returns>The conditional formatting object as a ContainsErrors rule</returns>
        public IExcelConditionalFormattingContainsErrors ContainsErrors
        {
            get
            {
                return _rule as IExcelConditionalFormattingContainsErrors;
            }
        }
        /// <summary>
        /// Returns the conditional formatting object as a ContainsText rule
        /// If this object is not of type ContainsText, null will be returned
        /// </summary>
        /// <returns>The conditional formatting object as a ContainsText rule</returns>
        public IExcelConditionalFormattingContainsText ContainsText
        {
            get
            {
                return _rule as IExcelConditionalFormattingContainsText;
            }
        }
        /// <summary>
        /// Returns the conditional formatting object as a NotContainsBlanks rule
        /// If this object is not of type NotContainsBlanks, null will be returned
        /// </summary>
        /// <returns>The conditional formatting object as a NotContainsBlanks rule</returns>
        public IExcelConditionalFormattingNotContainsBlanks NotContainsBlanks
        {
            get
            {
                return _rule as IExcelConditionalFormattingNotContainsBlanks;
            }
        }
        /// <summary>
        /// Returns the conditional formatting object as a NotContainsText rule
        /// If this object is not of type NotContainsText, null will be returned
        /// </summary>
        /// <returns>The conditional formatting object as a NotContainsText rule</returns>
        public IExcelConditionalFormattingNotContainsText NotContainsText
        {
            get
            {
                return _rule as IExcelConditionalFormattingNotContainsText;
            }
        }
        /// <summary>
        /// Returns the conditional formatting object as a NotContainsErrors rule
        /// If this object is not of type NotContainsErrors, null will be returned
        /// </summary>
        /// <returns>The conditional formatting object as a NotContainsErrors rule</returns>
        public IExcelConditionalFormattingNotContainsErrors NotContainsErrors
        {
            get
            {
                return _rule as IExcelConditionalFormattingNotContainsErrors;
            }
        }
        /// <summary>
        /// Returns the conditional formatting object as a NotBetween rule
        /// If this object is not of type NotBetween, null will be returned
        /// </summary>
        /// <returns>The conditional formatting object as a NotBetween rule</returns>
        public IExcelConditionalFormattingNotBetween NotBetween
        {
            get
            {
                return _rule as IExcelConditionalFormattingNotBetween;
            }
        }
        /// <summary>
        /// Returns the conditional formatting object as an Equal rule
        /// If this object is not of type Equal, null will be returned
        /// </summary>
        /// <returns>The conditional formatting object as an Equal rule</returns>
        public IExcelConditionalFormattingEqual Equal 
        { 
            get
            {
                return _rule as IExcelConditionalFormattingEqual;
            }
        }
        /// <summary>
        /// Returns the conditional formatting object as a NotEqual rule
        /// If this object is not of type NotEqual, null will be returned
        /// </summary>
        /// <returns>The conditional formatting object as a NotEqual rule</returns>
        public IExcelConditionalFormattingNotEqual NotEqual
        {
            get
            {
                return _rule as IExcelConditionalFormattingNotEqual;
            }   
        }
        /// <summary>
        /// Returns the conditional formatting object as a DuplicateValues rule
        /// If this object is not of type DuplicateValues, null will be returned
        /// </summary>
        /// <returns>The conditional formatting object as a DuplicateValues rule</returns>
        public IExcelConditionalFormattingDuplicateValues DuplicateValues
        {
            get
            {
                return _rule as IExcelConditionalFormattingDuplicateValues;
            }
        }
        /// <summary>
        /// Returns the conditional formatting object as a BeginsWith rule
        /// If this object is not of type BeginsWith, null will be returned
        /// </summary>
        /// <returns>The conditional formatting object as a BeginsWith rule</returns>
        public IExcelConditionalFormattingBeginsWith BeginsWith
        {
            get
            {
                return _rule as IExcelConditionalFormattingBeginsWith;
            }
        }
        /// <summary>
        /// Returns the conditional formatting object as an EndsWith rule
        /// If this object is not of type EndsWith, null will be returned
        /// </summary>
        /// <returns>The conditional formatting object as an EndsWith rule</returns>
        public IExcelConditionalFormattingEndsWith EndsWith
        {
            get
            {
                return _rule as IExcelConditionalFormattingEndsWith;
            }
        }
        /// <summary>
        /// Returns the conditional formatting object as an Expression rule
        /// If this object is not of type Expression, null will be returned
        /// </summary>
        /// <returns>The conditional formatting object as an Expression rule</returns>
        public IExcelConditionalFormattingExpression Expression
        {
            get
            {
                return _rule as IExcelConditionalFormattingExpression;
            }
        }
        /// <summary>
        /// Returns the conditional formatting object as a GreaterThan rule
        /// If this object is not of type GreaterThan, null will be returned
        /// </summary>
        /// <returns>The conditional formatting object as a GreaterThan rule</returns>
        public IExcelConditionalFormattingGreaterThan GreaterThan
        {
            get
            {
                return _rule as IExcelConditionalFormattingGreaterThan;
            }
        }
        /// <summary>
        /// Returns the conditional formatting object as a GreaterThanOrEqual rule
        /// If this object is not of type GreaterThanOrEqual, null will be returned
        /// </summary>
        /// <returns>The conditional formatting object as a GreaterThanOrEqual rule</returns>
        public IExcelConditionalFormattingGreaterThanOrEqual GreaterThanOrEqual
        {
            get
            {
                return _rule as IExcelConditionalFormattingGreaterThanOrEqual;
            }
        }
        /// <summary>
        /// Returns the conditional formatting object as a LessThan rule
        /// If this object is not of type LessThan, null will be returned
        /// </summary>
        /// <returns>The conditional formatting object as a LessThan rule</returns>
        public IExcelConditionalFormattingLessThan LessThan
        {
            get
            {
                return _rule as IExcelConditionalFormattingLessThan;
            }
        }
        /// <summary>
        /// Returns the conditional formatting object as a LessThanOrEqual rule
        /// If this object is not of type LessThanOrEqual, null will be returned
        /// </summary>
        /// <returns>The conditional formatting object as a LessThanOrEqual rule</returns>
        public IExcelConditionalFormattingLessThanOrEqual LessThanOrEqual
        {
            get
            {
                return _rule as IExcelConditionalFormattingLessThanOrEqual;
            }
        }
        /// <summary>
        /// Returns the conditional formatting object as a UniqueValues rule
        /// If this object is not of type UniqueValues, null will be returned
        /// </summary>
        /// <returns>The conditional formatting object as a UniqueValues rule</returns>
        public IExcelConditionalFormattingUniqueValues UniqueValues
        {
            get
            {
                return _rule as IExcelConditionalFormattingUniqueValues;
            }
        }
        /// <summary>
        /// Returns the conditional formatting object as a TwoColorScale rule
        /// If this object is not of type TwoColorScale, null will be returned
        /// </summary>
        /// <returns>The conditional formatting object as a TwoColorScale rule</returns>
        public IExcelConditionalFormattingTwoColorScale TwoColorScale
        {
            get
            {
                return _rule as IExcelConditionalFormattingTwoColorScale;
            }
        }
        /// <summary>
        /// Returns the conditional formatting object as a ThreeColorScale rule
        /// If this object is not of type ThreeColorScale, null will be returned
        /// </summary>
        /// <returns>The conditional formatting object as a ThreeColorScale rule</returns>
        public IExcelConditionalFormattingThreeColorScale ThreeColorScale
        {
            get
            {
                return _rule as IExcelConditionalFormattingThreeColorScale;
            }
        }
        /// <summary>
        /// Returns the conditional formatting object as a ThreeIconSet rule
        /// If this object is not of type ThreeIconSet, null will be returned
        /// </summary>
        /// <returns>The conditional formatting object as a ThreeIconSet rule</returns>
        public IExcelConditionalFormattingThreeIconSet<eExcelconditionalFormatting3IconsSetType> ThreeIconSet
        {
            get
            {
                return _rule as IExcelConditionalFormattingThreeIconSet<eExcelconditionalFormatting3IconsSetType>;
            }
        }
        /// <summary>
        /// Returns the conditional formatting object as a FourIconSet rule
        /// If this object is not of type FourIconSet, null will be returned
        /// </summary>
        /// <returns>The conditional formatting object as a FourIconSet rule</returns>
        public IExcelConditionalFormattingFourIconSet<eExcelconditionalFormatting4IconsSetType> FourIconSet
        {
            get
            {
                return _rule as IExcelConditionalFormattingFourIconSet<eExcelconditionalFormatting4IconsSetType>;
            }
        }
        /// <summary>
        /// Returns the conditional formatting object as a FiveIconSet rule
        /// If this object is not of type FiveIconSet, null will be returned
        /// </summary>
        /// <returns>The conditional formatting object as a FiveIconSet rule</returns>
        public IExcelConditionalFormattingFiveIconSet FiveIconSet
        {
            get
            {
                return _rule as IExcelConditionalFormattingFiveIconSet;
            }
        }
        /// <summary>
        /// Returns the conditional formatting object as a DataBar rule
        /// If this object is not of type DataBar, null will be returned
        /// </summary>
        /// <returns>The conditional formatting object as a DataBar rule</returns>
        public IExcelConditionalFormattingDataBarGroup DataBar
        {
            get
            {
                return _rule as IExcelConditionalFormattingDataBarGroup;
            }
        }
    }
}
