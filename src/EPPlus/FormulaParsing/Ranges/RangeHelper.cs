/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  03/03/2026         EPPlus Software AB           Virtual InMemoryRange support
 *************************************************************************************************/

namespace OfficeOpenXml.FormulaParsing.Ranges
{
    /// <summary>
    /// Helper methods for working with range physical/logical size.
    /// </summary>
    internal static class RangeHelper
    {
        /// <summary>
        /// Returns the effective number of data rows for a range.
        /// For virtual InMemoryRanges: the physical row count.
        /// For worksheet-backed ranges: constrained by worksheet dimension.
        /// For other ranges: Size.NumberOfRows.
        /// </summary>
        internal static int GetPhysicalRows(IRangeInfo range)
        {
            if (range is InMemoryRange imr)
                return imr.PhysicalRows;

            if (!range.IsInMemoryRange && range.Worksheet?.Dimension != null)
            {
                var adjusted = range.GetAddressDimensionAdjusted(0);
                if (adjusted != null)
                {
                    // If the adjusted address is invalid (no column overlap),
                    // the range column has no data at all
                    if (adjusted.FromCol > adjusted.ToCol || adjusted.FromRow > adjusted.ToRow)
                        return 0;
                    return adjusted.ToRow - adjusted.FromRow + 1;
                }
            }

            return range.Size.NumberOfRows;
        }
    }
}