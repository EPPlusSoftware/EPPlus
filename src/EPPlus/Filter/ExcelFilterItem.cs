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
using System.Globalization;

namespace OfficeOpenXml.Filter
{
    /// <summary>
    /// Base class for filter items
    /// </summary>
    public abstract class ExcelFilterItem
    {

    }
    /// <summary>
    /// A filter item for a value filter
    /// </summary>
    public class ExcelFilterValueItem : ExcelFilterItem
    {
        /// <summary>
        /// Inizialize the filter item
        /// </summary>
        /// <param name="value">The value to be filtered.</param>
        public ExcelFilterValueItem(string value)
        {
            Value = value;
            Utils.ConvertUtil.TryParseNumericString(value, out _valueDouble, CultureInfo.InvariantCulture);
            }
        /// <summary>
        /// A value to be filtered.
        /// </summary>
        public string Value { get; }
        internal double _valueDouble;
    }
}