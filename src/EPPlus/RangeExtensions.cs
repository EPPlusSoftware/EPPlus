/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  10/15/2023         EPPlus Software AB       Initial release EPPlus 7
 *************************************************************************************************/
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml
{
    /// <summary>
    /// Extension methods for <see cref="ExcelRangeBase"/>
    /// </summary>
    public static class RangeExtensions
    {
        /// <summary>
        /// Returns a new range, created by skipping a number of columns from the start.
        /// </summary>
        /// <param name="range">The source range</param>
        /// <param name="nColumns">The number of columns to skip</param>
        /// <returns>The result range</returns>
        public static ExcelRangeBase SkipColumns(this ExcelRangeBase range, int nColumns)
        {
            var nRangeColumns = range.End.Column - range.Start.Column + 1;
            if (nRangeColumns <= nColumns)
            {
                throw new IndexOutOfRangeException("SkipColumns: parameters nColumns must be less than number of columns in the source range");
            }
            var nRows = range.End.Row - range.Start.Row + 1;
            var cs = nColumns;
            var ce = range.End.Column - nColumns;
            return range.Offset(0, cs, nRows, ce);
        }

        /// <summary>
        /// Returns a new range, created by skipping a number of rows from the start.
        /// </summary>
        /// <param name="range">The source range</param>
        /// <param name="nRows">The number of rows to skip</param>
        /// <returns>The result range</returns>
        public static ExcelRangeBase SkipRows(this ExcelRangeBase range, int nRows)
        {
            var nRangeRows = range.End.Row - range.Start.Row + 1;
            if (nRangeRows <= nRows)
            {
                throw new IndexOutOfRangeException("SkipRows: parameters nRows must be less than number of columns in the source range");
            }
            var nCols = range.End.Column - range.Start.Column + 1;
            var rs = nRows;
            var re = range.End.Row - nRows;
            return range.Offset(rs, 0, re, nCols);
        }

        /// <summary>
        /// Returns a new range, created by taking a number of columns from the start.
        /// If <paramref name="nColumns"/> is greater than number of columns in the source range
        /// the entire source range will be returned.
        /// </summary>
        /// <param name="range">The source range</param>
        /// <param name="nColumns">The number of columns to take</param>
        /// <returns>The result range</returns>
        public static ExcelRangeBase TakeColumns(this ExcelRangeBase range, int nColumns)
        {
            var nRangeColumns = range.End.Column - range.Start.Column + 1;
            if (nRangeColumns <= nColumns)
            {
                return range;
            }
            var nRows = range.End.Row - range.Start.Row + 1;
            return range.Offset(0, 0, nRows, nColumns);
        }

        /// <summary>
        /// Returns a new range, created by taking a number of rows from the start.
        /// If <paramref name="nRows"/> is greater than number of rows in the source range
        /// the entire source range will be returned.
        /// </summary>
        /// <param name="range">The source range</param>
        /// <param name="nRows">The number of columns to take</param>
        /// <returns>The result range</returns>
        public static ExcelRangeBase TakeRows(this ExcelRangeBase range, int nRows)
        {
            var nRangeRows = range.End.Row - range.Start.Row + 1;
            if (nRangeRows <= nRows)
            {
                return range;
            }
            var nCols = range.End.Column - range.Start.Column + 1;
            return range.Offset(0, 0, nRows, nCols);
        }
    }
}
