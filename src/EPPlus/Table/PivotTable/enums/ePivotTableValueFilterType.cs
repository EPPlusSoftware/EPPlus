/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  09/02/2020         EPPlus Software AB       EPPlus 5.4
 *************************************************************************************************/
namespace OfficeOpenXml.Table.PivotTable
{
    /// <summary>
    /// Defines a pivot table value filter type for numbers and strings
    /// </summary>
    public enum ePivotTableValueFilterType
    {
        /// <summary>
        /// A numeric or string filter - Value Between
        /// </summary>
        ValueBetween = ePivotTableFilterType.ValueBetween,
        /// <summary>
        /// A numeric or string filter - Equal
        /// </summary>
        ValueEqual = ePivotTableFilterType.ValueEqual,
        /// <summary>
        /// A numeric or string filter - GreaterThan
        /// </summary>
        ValueGreaterThan = ePivotTableFilterType.ValueGreaterThan,
        /// <summary>
        /// A numeric or string filter - Greater Than Or Equal
        /// </summary>
        ValueGreaterThanOrEqual = ePivotTableFilterType.ValueGreaterThanOrEqual,
        /// <summary>
        /// A numeric or string filter - Less Than 
        /// </summary>
        ValueLessThan = ePivotTableFilterType.ValueLessThan,
        /// <summary>
        /// A numeric or string filter - Less Than Or Equal
        /// </summary>
        ValueLessThanOrEqual = ePivotTableFilterType.ValueLessThanOrEqual,
        /// <summary>
        /// A numeric or string filter - Not Between
        /// </summary>
        ValueNotBetween = ePivotTableFilterType.ValueNotBetween,
        /// <summary>
        /// A numeric or string filter - Not Equal
        /// </summary>
        ValueNotEqual = ePivotTableFilterType.ValueNotEqual
    }
}
