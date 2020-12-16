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
    [AttributeUsage(AttributeTargets.Class | AttributeTargets.Interface)]
    public class EpplusTableAttribute : Attribute
    {
        /// <summary>
        /// Constructor
        /// </summary>
        public EpplusTableAttribute()
        {
            TableStyle = TableStyles.None;
        }
        /// <summary>
        /// Table style
        /// </summary>
        public TableStyles TableStyle
        {
            get;
            set;
        }

        /// <summary>
        /// If true, there will be a header row with column names over the data
        /// </summary>
        public bool PrintHeaders
        {
            get;
            set;
        }

        /// <summary>
        /// If true, the first column of the table is highlighted
        /// </summary>
        public bool ShowFirstColumn
        {
            get;
            set;
        }

        /// <summary>
        /// If true, the last column of the table is highlighted
        /// </summary>
        public bool ShowLastColumn
        {
            get;
            set;
        }

        /// <summary>
        /// If true, a totals row will be added under the table data. This should be used in combination with <see cref="EpplusTableColumnAttributeBase.TotalsRowFunction"/> on the column attributes.
        /// </summary>
        public bool ShowTotal
        {
            get;
            set;
        }

        /// <summary>
        /// If true, column width will be adjusted to cell content
        /// </summary>
        public bool AutofitColumns
        {
            get;
            set;
        }

        /// <summary>
        /// If true, EPPlus will calculate the table range when the data has been read into the spreadsheet and store the results
        /// in the Value property of each cell.
        /// </summary>
        public bool AutoCalculate
        {
            get;
            set;
        }
    }
}
