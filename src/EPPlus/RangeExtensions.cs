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
        /// <param name="count">The number of columns to skip</param>
        /// <returns>The result range</returns>
        public static ExcelRangeBase SkipColumns(this ExcelRangeBase range, int count)
        {
            var nRangeColumns = range.End.Column - range.Start.Column + 1;
            if (nRangeColumns <= count)
            {
                throw new IndexOutOfRangeException("SkipColumns: parameters nColumns must be less than number of columns in the source range");
            }
            var nRows = range.End.Row - range.Start.Row + 1;
            var cs = count;
            var ce = range.End.Column - count;
            return range.Offset(0, cs, nRows, ce);
        }

        /// <summary>
        /// Returns a new range, created by skipping a number of rows from the start.
        /// </summary>
        /// <param name="range">The source range</param>
        /// <param name="count">The number of rows to skip</param>
        /// <returns>The result range</returns>
        public static ExcelRangeBase SkipRows(this ExcelRangeBase range, int count)
        {
            var nRangeRows = range.End.Row - range.Start.Row + 1;
            if (nRangeRows <= count)
            {
                throw new IndexOutOfRangeException("SkipRows: parameters nRows must be less than number of columns in the source range");
            }
            var nCols = range.End.Column - range.Start.Column + 1;
            var rs = count;
            var re = range.End.Row - count;
            return range.Offset(rs, 0, re, nCols);
        }

        /// <summary>
        /// Returns a new range, created by taking a number of columns from the start.
        /// If <paramref name="count"/> is greater than number of columns in the source range
        /// the entire source range will be returned.
        /// </summary>
        /// <param name="range">The source range</param>
        /// <param name="count">The number of columns to take</param>
        /// <returns>The result range</returns>
        public static ExcelRangeBase TakeColumns(this ExcelRangeBase range, int count)
        {
            var nRangeColumns = range.End.Column - range.Start.Column + 1;
            if (nRangeColumns <= count)
            {
                return range;
            }
            var nRows = range.End.Row - range.Start.Row + 1;
            return range.Offset(0, 0, nRows, count);
        }

        /// <summary>
        /// Returns a new range, created by taking a number of rows from the start.
        /// If <paramref name="count"/> is greater than number of rows in the source range
        /// the entire source range will be returned.
        /// </summary>
        /// <param name="range">The source range</param>
        /// <param name="count">The number of columns to take</param>
        /// <returns>The result range</returns>
        public static ExcelRangeBase TakeRows(this ExcelRangeBase range, int count)
        {
            var nRangeRows = range.End.Row - range.Start.Row + 1;
            if (nRangeRows <= count)
            {
                return range;
            }
            var nCols = range.End.Column - range.Start.Column + 1;
            return range.Offset(0, 0, count, nCols);
        }

        /// <summary>
        /// Returns a new range, created by taking a specific number of columns between from the offset parameter.
        /// </summary>
        /// <param name="range">The source range</param>
        /// <param name="offset">Offset of the start-column (zero-based)</param>
        /// <param name="count">The number of columns to take</param>
        /// <returns>The result range</returns>
        public static ExcelRangeBase TakeColumnsBetween(this ExcelRangeBase range, int offset, int count = 1)
        {
            var nRangeColumns = range.End.Column - range.Start.Column + 1;
            if(offset >= nRangeColumns)
            {
                throw new ArgumentException("Parameter offset must be less than number of columns in the range");
            }
            else if (offset < 0)
            {
                throw new ArgumentException("Parameter offset cannot be a negative number.");
            }
            else if(offset + count > nRangeColumns)
            {
                throw new IndexOutOfRangeException("offset + count cannot be larger than number of columns in the range");
            }
            var nRows = range.End.Row - range.Start.Row + 1;
            return range.Offset(0, offset, nRows, count);
        }

        /// <summary>
        /// Returns a new range, created by taking a specific number of rows based on the offset parameter.
        /// </summary>
        /// <param name="range">The source range</param>
        /// <param name="offset">Offset of the start-row (zero-based)</param>
        /// <param name="count">The number of rows to take</param>
        /// <returns>The result range</returns>
        public static ExcelRangeBase TakeRowsBetween(this ExcelRangeBase range, int offset, int count = 1)
        {
            var nRangeRows = range.End.Row - range.Start.Row + 1;
            if (offset >= nRangeRows)
            {
                throw new ArgumentException("Parameter offset must be less than number of rows in the range");
            }
            else if (offset < 0)
            {
                throw new ArgumentException("Parameter offset cannot be a negative number.");
            }
            else if (offset + count > nRangeRows)
            {
                throw new IndexOutOfRangeException("offset + count cannot be larger than number of rows in the range");
            }
            var nCols = range.End.Column - range.Start.Column + 1;
            return range.Offset(offset, 0, count, nCols);
        }
    }
}
