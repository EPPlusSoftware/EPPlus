/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  12/10/2020         EPPlus Software AB       EPPlus 5.5
 *************************************************************************************************/
using OfficeOpenXml.Table;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Attributes
{
    /// <summary>
    /// Attribute used by <see cref="ExcelRangeBase.LoadFromCollection{T}(IEnumerable{T})" /> to configure column parameters for the functions/>
    /// </summary>
    public abstract class EpplusTableColumnAttributeBase : Attribute
    {

        /// <summary>
        /// Order of the columns value, default value is 0
        /// </summary>
        public int Order
        {
            get;
            set;
        } = int.MaxValue;

        /// <summary>
        /// Name shown in the header row, overriding the property name
        /// </summary>
        public string Header
        {
            get;
            set;
        }

        /// <summary>
        /// Excel format string for the column
        /// </summary>
        public string NumberFormat
        {
            get;
            set;
        }

        /// <summary>
        /// A number to be used in a NumberFormatProvider.
        /// Default value is int.MinValue, which means it will be ignored.
        /// </summary>
        public int NumberFormatId
        {
            get;
            set;
        } = int.MinValue;

        /// <summary>
        /// If true, the entire column will be hidden.
        /// </summary>
        public bool Hidden
        {
            get;
            set;
        }

        /// <summary>
        /// Indicates whether the Built in (default) hyperlink style should be
        /// applied to hyperlinks or not. Default value is true.
        /// </summary>
        public bool UseBuiltInHyperlinkStyle
        {
            get; set;
        } = true;

        /// <summary>
        /// If not <see cref="RowFunctions.None"/> the last cell in the column (the totals row) will contain a formula of the specified type.
        /// </summary>
        public RowFunctions TotalsRowFunction
        {
            get;
            set;
        } = RowFunctions.None;

        /// <summary>
        /// Formula for the total row of this column.
        /// </summary>
        public string TotalsRowFormula
        {
            get;
            set;
        }

        /// <summary>
        /// Number format for this columns cell in the totals row.
        /// </summary>
        public string TotalsRowNumberFormat
        {
            get;
            set;
        }

        /// <summary>
        /// String in this columns cell in the totals row
        /// </summary>
        public string TotalsRowLabel
        {
            get;
            set;
        }
    }
}
