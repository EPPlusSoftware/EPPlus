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

namespace OfficeOpenXml
{
    /// <summary>
    /// Extended address information for a table address
    /// </summary>
    public class ExcelTableAddress
    {
        /// <summary>
        /// The name of the table
        /// </summary>
        public string Name { get; set; }
        /// <summary>
        /// Column span
        /// </summary>
        public string ColumnSpan { get; set; }
        /// <summary>
        /// Reference entire table
        /// </summary>
        public bool IsAll { get; set; }
        /// <summary>
        /// Reference the table header row
        /// </summary>
        public bool IsHeader { get; set; }
        /// <summary>
        /// Reference table data
        /// </summary>
        public bool IsData { get; set; }
        /// <summary>
        /// Reference table totals row
        /// </summary>
        public bool IsTotals { get; set; }
        /// <summary>
        /// Reference the current table row
        /// </summary>
        public bool IsThisRow { get; set; }
    }
}
