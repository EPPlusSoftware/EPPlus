/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  05/08/2026         EPPlus Software AB       Initial release
 *************************************************************************************************/
using OfficeOpenXml.FormulaParsing.Excel.Functions.MathFunctions;

namespace OfficeOpenXml.Utils.Drawings
{
    /// <summary>
    /// Helper methods for converting between worksheet coordinates and pixels.
    /// </summary>
    internal static class PixelHelper
    {
        /// <summary>
        /// Returns the width of a column in pixels, using the same formula
        /// Excel uses internally.
        /// </summary>
        /// <param name="ws">The worksheet.</param>
        /// <param name="column">The 1-based column index.</param>
        /// <returns>The column width in pixels.</returns>
        internal static double GetColumnWidth(ExcelWorksheet ws, int column)
        {
            double mdw = ws.Workbook.MaxFontWidth;
            return MathHelper.TruncateDouble(
                ((256 * ws.GetColumnWidth(column) + MathHelper.TruncateDouble(128 / mdw)) / 256) * mdw);
        }

        /// <summary>
        /// Returns the height of a row in pixels.
        /// </summary>
        /// <param name="ws">The worksheet.</param>
        /// <param name="row">The 1-based row index.</param>
        /// <returns>The row height in pixels.</returns>
        internal static double GetRowHeight(ExcelWorksheet ws, int row)
        {
            return ws.GetRowHeight(row) / 0.75;
        }
    }
}