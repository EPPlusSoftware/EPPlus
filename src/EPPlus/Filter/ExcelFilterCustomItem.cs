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
namespace OfficeOpenXml.Filter
{
    /// <summary>
    /// A custom filter item
    /// </summary>
    public class ExcelFilterCustomItem : ExcelFilterValueItem
    {
        /// <summary>
        /// Create a Custom filter.
        /// </summary>
        /// <param name="value">The value to filter by. 
        /// If the data is text wildcard can be used. 
        /// Asterisk (*) for any combination of characters. 
        /// Question mark (?) for any single charcter
        /// If the data is numeric, use dot (.) for decimal.</param>
        /// <param name="filterOperator">The operator to use</param>
        public ExcelFilterCustomItem(string value, eFilterOperator filterOperator = eFilterOperator.Equal) : base(value)
        {
            Operator = filterOperator;
        }
        /// <summary>
        /// Operator used by the filter comparison
        /// </summary>
        public eFilterOperator? Operator { get; set; }
    }
}