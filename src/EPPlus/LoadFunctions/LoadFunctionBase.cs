/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  07/16/2020         EPPlus Software AB       EPPlus 5.2.1
 *************************************************************************************************/
using OfficeOpenXml.Table;
using System;
using System.Collections.Generic;
using System.Text;

namespace OfficeOpenXml.LoadFunctions
{
    /// <summary>
    /// Base class for ExcelRangeBase.LoadFrom[...] functions
    /// </summary>
    internal abstract class LoadFunctionBase
    {
        public LoadFunctionBase(ExcelRangeBase range, bool printHeaders, TableStyles tableStyle)
        {
            Range = range;
            PrintHeaders = printHeaders;
            TableStyle = tableStyle;
        }

        /// <summary>
        /// The range to which the data should be loaded
        /// </summary>
        protected ExcelRangeBase Range { get; }

        /// <summary>
        /// If true a header row will be printed above the data
        /// </summary>
        protected bool PrintHeaders { get; }

        /// <summary>
        /// If value is other than TableStyles.None the data will be added to a table in the worksheet.
        /// </summary>
        protected TableStyles TableStyle { get; set; }

        /// <summary>
        /// Returns how many rows there are in the range (header row not included)
        /// </summary>
        /// <returns></returns>
        protected abstract int GetNumberOfRows();

        /// <summary>
        /// Returns how many columns there are in the range
        /// </summary>
        /// <returns></returns>
        protected abstract int GetNumberOfColumns();

        protected abstract void LoadInternal(object[,] values);

        /// <summary>
        /// Loads the data into the worksheet
        /// </summary>
        /// <returns></returns>
        internal ExcelRangeBase Load()
        {
            var nRows = PrintHeaders ? GetNumberOfRows() + 1 : GetNumberOfRows();
            var nCols = GetNumberOfColumns();
            var values = new object[nRows, nCols];
            LoadInternal(values);
            var ws = Range.Worksheet;
            ws.SetRangeValueInner(Range._fromRow, Range._fromCol, Range._fromRow + nRows - 1, Range._fromCol + nCols - 1, values);
            
            //Must have at least 1 row, if header is shown
            if (nRows == 1 && PrintHeaders)
            {
                nRows++;
            }

            var r = ws.Cells[Range._fromRow, Range._fromCol, Range._fromRow + nRows - 1, Range._fromCol + nCols - 1];

            if (TableStyle != TableStyles.None)
            {
                var tbl = ws.Tables.Add(r, "");
                tbl.ShowHeader = PrintHeaders;
                tbl.TableStyle = TableStyle;
            }

            return r;
        }
    }
}
