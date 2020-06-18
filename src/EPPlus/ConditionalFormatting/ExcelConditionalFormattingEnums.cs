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

namespace OfficeOpenXml.ConditionalFormatting
{
  /// <summary>
  /// Enum for Conditional Format Type ST_CfType §18.18.12. With some changes.
  /// </summary>
  public enum eExcelConditionalFormattingRuleType
  {
    #region Average
    /// <summary>
    /// Highlights cells that are above the average for all values in the range.
    /// </summary>
    /// <remarks>AboveAverage Excel CF Rule Type</remarks>
    AboveAverage,

    /// <summary>
    /// Highlights cells that are above or equal to the average for all values in the range.
    /// </summary>
    /// <remarks>AboveAverage Excel CF Rule Type</remarks>
    AboveOrEqualAverage,

    /// <summary>
    /// Highlights cells that are below the average for all values in the range.
    /// </summary>
    /// <remarks>AboveAverage Excel CF Rule Type</remarks>
    BelowAverage,

    /// <summary>
    /// Highlights cells that are below or equal to the average for all values in the range.
    /// </summary>
    /// <remarks>AboveAverage Excel CF Rule Type</remarks>
    BelowOrEqualAverage,
        #endregion

    #region StdDev
    /// <summary>
    /// Highlights cells that are above the standard deviation for all values in the range.
    /// <remarks>AboveAverage Excel CF Rule Type</remarks>
    /// </summary>
    AboveStdDev,

    /// <summary>
    /// Highlights cells that are below the standard deviation for all values in the range.
    /// </summary>
    /// <remarks>AboveAverage Excel CF Rule Type</remarks>
    BelowStdDev,
    #endregion

    #region TopBottom
    /// <summary>
    /// Highlights cells whose values fall in the bottom N bracket as specified.
    /// </summary>
    /// <remarks>Top10 Excel CF Rule Type</remarks>
    Bottom,

    /// <summary>
    /// Highlights cells whose values fall in the bottom N percent as specified.
    /// </summary>
    /// <remarks>Top10 Excel CF Rule Type</remarks>
    BottomPercent,

    /// <summary>
    /// Highlights cells whose values fall in the top N bracket as specified.
    /// </summary>
    /// <remarks>Top10 Excel CF Rule Type</remarks>
    Top,

    /// <summary>
    /// Highlights cells whose values fall in the top N percent as specified.
    /// </summary>
    /// <remarks>Top10 Excel CF Rule Type</remarks>
    TopPercent,
    #endregion

    #region TimePeriod
    /// <summary>
    /// Highlights cells containing dates in the last 7 days.
    /// </summary>
    /// <remarks>TimePeriod Excel CF Rule Type</remarks>
    Last7Days,

    /// <summary>
    /// Highlights cells containing dates in the last month.
    /// </summary>
    /// <remarks>TimePeriod Excel CF Rule Type</remarks>
    LastMonth,

    /// <summary>
    /// Highlights cells containing dates in the last week.
    /// </summary>
    /// <remarks>TimePeriod Excel CF Rule Type</remarks>
    LastWeek,

    /// <summary>
    /// Highlights cells containing dates in the next month.
    /// </summary>
    /// <remarks>TimePeriod Excel CF Rule Type</remarks>
    NextMonth,

    /// <summary>
    /// Highlights cells containing dates in the next week.
    /// </summary>
    /// <remarks>TimePeriod Excel CF Rule Type</remarks>
    NextWeek,

    /// <summary>
    /// Highlights cells containing dates in this month.
    /// </summary>
    /// <remarks>TimePeriod Excel CF Rule Type</remarks>
    ThisMonth,

    /// <summary>
    /// Highlights cells containing dates in this week.
    /// </summary>
    /// <remarks>TimePeriod Excel CF Rule Type</remarks>
    ThisWeek,

    /// <summary>
    /// Highlights cells containing todays date.
    /// </summary>
    /// <remarks>TimePeriod Excel CF Rule Type</remarks>
    Today,

    /// <summary>
    /// Highlights cells containing tomorrows date.
    /// </summary>
    /// <remarks>TimePeriod Excel CF Rule Type</remarks>
    Tomorrow,

    /// <summary>
    /// Highlights cells containing yesterdays date.
    /// </summary>
    /// <remarks>TimePeriod Excel CF Rule Type</remarks>
    Yesterday,
    #endregion

    #region CellIs
    /// <summary>
    /// Highlights cells in the range that begin with the given text.
    /// </summary>
    /// <remarks>
    /// Equivalent to using the LEFT() sheet function and comparing values.
    /// </remarks>
    /// <remarks>BeginsWith Excel CF Rule Type</remarks>
    BeginsWith,

    /// <summary>
    /// Highlights cells in the range between the given two formulas.
    /// </summary>
    /// <remarks>CellIs Excel CF Rule Type</remarks>
    Between,

    /// <summary>
    /// Highlights cells that are completely blank.
    /// </summary>
    /// <remarks>
    /// Equivalent of using LEN(TRIM()). This means that if the cell contains only
    /// characters that TRIM() would remove, then it is considered blank. An empty cell
    /// is also considered blank.
    /// </remarks>
    /// <remarks>ContainsBlanks Excel CF Rule Type</remarks>
    ContainsBlanks,

    /// <summary>
    /// Highlights cells with formula errors.
    /// </summary>
    /// <remarks>
    /// Equivalent to using ISERROR() sheet function to determine if there is
    /// a formula error.
    /// </remarks>
    /// <remarks>ContainsErrors Excel CF Rule Type</remarks>
    ContainsErrors,

    /// <summary>
    /// Highlights cells in the range that begin with
    /// the given text.
    /// </summary>
    /// <remarks>
    /// Equivalent to using the LEFT() sheet function and comparing values.
    /// </remarks>
    /// <remarks>ContainsText Excel CF Rule Type</remarks>
    ContainsText,

    /// <summary>
    /// Highlights duplicated values.
    /// </summary>
    /// <remarks>DuplicateValues Excel CF Rule Type</remarks>
    DuplicateValues,

    /// <summary>
    /// Highlights cells ending with the given text.
    /// </summary>
    /// <remarks>
    /// Equivalent to using the RIGHT() sheet function and comparing values.
    /// </remarks>
    /// <remarks>EndsWith Excel CF Rule Type</remarks>
    EndsWith,

    /// <summary>
    /// Highlights cells equal to the given formula.
    /// </summary>
    /// <remarks>CellIs Excel CF Rule Type</remarks>
    Equal,

    /// <summary>
    /// This rule contains a formula to evaluate. When the formula result is true, the cell is highlighted.
    /// </summary>
    /// <remarks>Expression Excel CF Rule Type</remarks>
    Expression,

    /// <summary>
    /// Highlights cells greater than the given formula.
    /// </summary>
    /// <remarks>CellIs Excel CF Rule Type</remarks>
    GreaterThan,

    /// <summary>
    /// Highlights cells greater than or equal the given formula.
    /// </summary>
    /// <remarks>CellIs Excel CF Rule Type</remarks>
    GreaterThanOrEqual,

    /// <summary>
    /// Highlights cells less than the given formula.
    /// </summary>
    /// <remarks>CellIs Excel CF Rule Type</remarks>
    LessThan,

    /// <summary>
    /// Highlights cells less than or equal the given formula.
    /// </summary>
    /// <remarks>CellIs Excel CF Rule Type</remarks>
    LessThanOrEqual,

    /// <summary>
    /// Highlights cells outside the range in given two formulas.
    /// </summary>
    /// <remarks>CellIs Excel CF Rule Type</remarks>
    NotBetween,

    /// <summary>
    /// Highlights cells that does not contains the given formula.
    /// </summary>
    /// <remarks>CellIs Excel CF Rule Type</remarks>
    NotContains,

    /// <summary>
    /// Highlights cells that are not blank.
    /// </summary>
    /// <remarks>
    /// Equivalent of using LEN(TRIM()). This means that if the cell contains only
    /// characters that TRIM() would remove, then it is considered blank. An empty cell
    /// is also considered blank.
    /// </remarks>
    /// <remarks>NotContainsBlanks Excel CF Rule Type</remarks>
    NotContainsBlanks,

    /// <summary>
    /// Highlights cells without formula errors.
    /// </summary>
    /// <remarks>
    /// Equivalent to using ISERROR() sheet function to determine if there is a
    /// formula error.
    /// </remarks>
    /// <remarks>NotContainsErrors Excel CF Rule Type</remarks>
    NotContainsErrors,

    /// <summary>
    /// Highlights cells that do not contain the given text.
    /// </summary>
    /// <remarks>
    /// Equivalent to using the SEARCH() sheet function.
    /// </remarks>
    /// <remarks>NotContainsText Excel CF Rule Type</remarks>
    NotContainsText,

    /// <summary>
    ///     .
    /// </summary>
    /// <remarks>CellIs Excel CF Rule Type</remarks>
    NotEqual,

    /// <summary>
    /// Highlights unique values in the range.
    /// </summary>
    /// <remarks>UniqueValues Excel CF Rule Type</remarks>
    UniqueValues,
    #endregion

    #region ColorScale
    /// <summary>
    /// Three Color Scale (Low, Middle and High Color Scale)
    /// </summary>
    /// <remarks>ColorScale Excel CF Rule Type</remarks>
    ThreeColorScale,

    /// <summary>
    /// Two Color Scale (Low and High Color Scale)
    /// </summary>
    /// <remarks>ColorScale Excel CF Rule Type</remarks>
    TwoColorScale,
    #endregion

    #region IconSet
    /// <summary>
    /// This conditional formatting rule applies a 3 set icons to cells according
    /// to their values.
    /// </summary>
    /// <remarks>IconSet Excel CF Rule Type</remarks>
    ThreeIconSet,

    /// <summary>
    /// This conditional formatting rule applies a 4 set icons to cells according
    /// to their values.
    /// </summary>
    /// <remarks>IconSet Excel CF Rule Type</remarks>
    FourIconSet,

    /// <summary>
    /// This conditional formatting rule applies a 5 set icons to cells according
    /// to their values.
    /// </summary>
    /// <remarks>IconSet Excel CF Rule Type</remarks>
    FiveIconSet,
    #endregion

    #region DataBar
    /// <summary>
    /// This conditional formatting rule displays a gradated data bar in the range of cells.
    /// </summary>
    /// <remarks>DataBar Excel CF Rule Type</remarks>
    DataBar
    #endregion
  }

  /// <summary>
  /// Enum for Conditional Format Value Object Type ST_CfvoType §18.18.13
  /// </summary>
  public enum eExcelConditionalFormattingValueObjectType
  {
    /// <summary>
    /// Formula
    /// </summary>
    Formula,

    /// <summary>
    /// Maximum Value
    /// </summary>
    Max,

    /// <summary>
    /// Minimum Value
    /// </summary>
    Min,

    /// <summary>
    /// Number Value
    /// </summary>
    Num,

    /// <summary>
    /// Percent
    /// </summary>
    Percent,

    /// <summary>
    /// Percentile
    /// </summary>
    Percentile
  }

  /// <summary>
  /// Enum for Conditional Formatting Value Object Position
  /// </summary>
  public enum eExcelConditionalFormattingValueObjectPosition
  {
    /// <summary>
    /// The lower position for both TwoColorScale and ThreeColorScale
    /// </summary>
    Low,

    /// <summary>
    /// The middle position only for ThreeColorScale
    /// </summary>
    Middle,

    /// <summary>
    /// The highest position for both TwoColorScale and ThreeColorScale
    /// </summary>
    High
  }

  /// <summary>
  /// Enum for Conditional Formatting Value Object Node Type
  /// </summary>
  public enum eExcelConditionalFormattingValueObjectNodeType
  {
    /// <summary>
    /// 'cfvo' node
    /// </summary>
    Cfvo,

    /// <summary>
    /// 'color' node
    /// </summary>
    Color
  }

  /// <summary>
  /// Enum for Conditional Formatting Operartor Type ST_ConditionalFormattingOperator §18.18.15
  /// </summary>
  public enum eExcelConditionalFormattingOperatorType
  {
    /// <summary>
    /// Begins With. 'Begins with' operator
    /// </summary>
    BeginsWith,

    /// <summary>
    /// Between. 'Between' operator
    /// </summary>
    Between,

    /// <summary>
    /// Contains. 'Contains' operator
    /// </summary>
    ContainsText,

    /// <summary>
    /// Ends With. 'Ends with' operator
    /// </summary>
    EndsWith,

    /// <summary>
    /// Equal. 'Equal to' operator
    /// </summary>
    Equal,

    /// <summary>
    /// Greater Than. 'Greater than' operator
    /// </summary>
    GreaterThan,

    /// <summary>
    /// Greater Than Or Equal. 'Greater than or equal to' operator
    /// </summary>
    GreaterThanOrEqual,

    /// <summary>
    /// Less Than. 'Less than' operator
    /// </summary>
    LessThan,

    /// <summary>
    /// Less Than Or Equal. 'Less than or equal to' operator
    /// </summary>
    LessThanOrEqual,

    /// <summary>
    /// Not Between. 'Not between' operator
    /// </summary>
    NotBetween,

    /// <summary>
    /// Does Not Contain. 'Does not contain' operator
    /// </summary>
    NotContains,

    /// <summary>
    /// Not Equal. 'Not equal to' operator
    /// </summary>
    NotEqual
  }

  /// <summary>
  /// Enum for Conditional Formatting Time Period Type ST_TimePeriod §18.18.82
  /// </summary>
  public enum eExcelConditionalFormattingTimePeriodType
  {
    /// <summary>
    /// Last 7 Days. A date in the last seven days.
    /// </summary>
    Last7Days,

    /// <summary>
    /// Last Month. A date occuring in the last calendar month.
    /// </summary>
    LastMonth,

    /// <summary>
    /// Last Week. A date occuring last week.
    /// </summary>
    LastWeek,

    /// <summary>
    /// Next Month. A date occuring in the next calendar month.
    /// </summary>
    NextMonth,

    /// <summary>
    /// Next Week. A date occuring next week.
    /// </summary>
    NextWeek,

    /// <summary>
    /// This Month. A date occuring in this calendar month.
    /// </summary>
    ThisMonth,

    /// <summary>
    /// This Week. A date occuring this week.
    /// </summary>
    ThisWeek,

    /// <summary>
    /// Today. Today's date.
    /// </summary>
    Today,

    /// <summary>
    /// Tomorrow. Tomorrow's date.
    /// </summary>
    Tomorrow,

    /// <summary>
    /// Yesterday. Yesterday's date.
    /// </summary>
    Yesterday
  }

  /// <summary>
  /// 18.18.42 ST_IconSetType (Icon Set Type) - Only 3 icons
  /// </summary>
  public enum eExcelconditionalFormatting3IconsSetType
  {
    /// <summary>
    /// 3 arrows icon set.
    /// </summary>
    Arrows,

    /// <summary>
    /// 3 gray arrows icon set.
    /// </summary>
    ArrowsGray,

    /// <summary>
    /// 3 flags icon set. 
    /// </summary>
    Flags,

    /// <summary>
    /// 3 signs icon set.
    /// </summary>
    Signs,

    /// <summary>
    /// 3 symbols icon set.
    /// </summary>
    Symbols,

    /// <summary>
    /// 3 Symbols icon set.
    /// </summary>
    Symbols2,

    /// <summary>
    /// 3 traffic lights icon set (#1).
    /// </summary>
    TrafficLights1,

    /// <summary>
    /// 3 traffic lights icon set with thick black border.
    /// </summary>
    TrafficLights2
  }

  /// <summary>
  /// 18.18.42 ST_IconSetType (Icon Set Type) - Only 4 icons
  /// </summary>
  public enum eExcelconditionalFormatting4IconsSetType
  {
    /// <summary>
    /// (4 Arrows) 4 arrows icon set.
    /// </summary>
    Arrows,

    /// <summary>
    /// (4 Arrows (Gray)) 4 gray arrows icon set.
    /// </summary>
    ArrowsGray,

    /// <summary>
    /// (4 Ratings) 4 ratings icon set.
    /// </summary>
    Rating,

    /// <summary>
    /// (4 Red To Black) 4 'red to black' icon set.
    /// </summary>
    RedToBlack,

    /// <summary>
    /// (4 Traffic Lights) 4 traffic lights icon set.
    /// </summary>
    TrafficLights
  }

  /// <summary>
  /// 18.18.42 ST_IconSetType (Icon Set Type) - Only 5 icons
  /// </summary>
  public enum eExcelconditionalFormatting5IconsSetType
  {
    /// <summary>
    /// 5 arrows icon set.
    /// </summary>
    Arrows,

    /// <summary>
    /// 5 gray arrows icon set.
    /// </summary>
    ArrowsGray,

    /// <summary>
    /// 5 quarters icon set.
    /// </summary>
    Quarters,

    /// <summary>
    /// 5 rating icon set.
    /// </summary>
    Rating
  }
  /// <summary>
  /// 18.18.42 ST_IconSetType (Icon Set Type)
  /// </summary>
  public enum eExcelconditionalFormattingIconsSetType
  {
      /// <summary>
      /// 3 arrows icon set
      /// </summary>
      ThreeArrows,

      /// <summary>
      /// 3 gray arrows icon set
      /// </summary>
      ThreeArrowsGray,

      /// <summary>
      /// 3 flags icon set. 
      /// </summary>
      ThreeFlags,

      /// <summary>
      /// 3 signs icon set.
      /// </summary>
      ThreeSigns,

      /// <summary>
      /// 3 symbols icon set.
      /// </summary>
      ThreeSymbols,

      /// <summary>
      /// 3 Symbols icon set.
      /// </summary>
      ThreeSymbols2,

      /// <summary>
      /// 3 traffic lights icon set (#1).
      /// </summary>
      ThreeTrafficLights1,

      /// <summary>
      /// 3 traffic lights icon set with thick black border.
      /// </summary>
      ThreeTrafficLights2,
 
    /// <summary>
    /// 4 arrows icon set.
    /// </summary>
    FourArrows,

    /// <summary>
    /// 4 gray arrows icon set.
    /// </summary>
    FourArrowsGray,

    /// <summary>
    /// 4 ratings icon set.
    /// </summary>
    FourRating,

    /// <summary>
    /// 4 'red to black' icon set.
    /// </summary>
    FourRedToBlack,

    /// <summary>
    /// 4 traffic lights icon set.
    /// </summary>
    FourTrafficLights,

    /// <summary>
    /// 5 arrows icon set.
    /// </summary>
    FiveArrows,

    /// <summary>
    /// 5 gray arrows icon set.
    /// </summary>
    FiveArrowsGray,

    /// <summary>
    /// 5 quarters icon set.
    /// </summary>
    FiveQuarters,

    /// <summary>
    /// 5 rating icon set.
    /// </summary>
    FiveRating
}
}