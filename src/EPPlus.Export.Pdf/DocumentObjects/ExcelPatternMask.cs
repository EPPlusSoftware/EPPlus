/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  27/11/2025         EPPlus Software AB           EPPlus 9
 *************************************************************************************************/
using System.Collections.Generic;

namespace EPPlus.Export.Pdf.DocumentObjects
{
    /// <summary>
    /// The cell fill pattern types EPPlus renders in the PDF export, matching
    /// the names of OfficeOpenXml.Style.ExcelFillStyle. The pattern geometry is
    /// taken from how Microsoft Excel itself rasterises each pattern when it
    /// exports to PDF (an 8x8 tile), NOT from the ECMA-376 ST_Shd masks. Excel
    /// and ST_Shd only overlap for a few patterns; since the goal is visual
    /// parity with Excel, Excel's own output is the reference.
    /// </summary>
    internal enum ExcelPatternMask
    {
        DarkGray,
        MediumGray,
        LightGray,
        Gray125,
        Gray0625,
        DarkHorizontal,
        DarkVertical,
        DarkDown,
        DarkUp,
        DarkGrid,
        DarkTrellis,
        LightHorizontal,
        LightVertical,
        LightDown,
        LightUp,
        LightGrid,
        LightTrellis,
    }

    /// <summary>
    /// Reference catalog of the 8x8 cell fill pattern masks, transcribed from the
    /// bitmaps Microsoft Excel produces when exporting each pattern fill to PDF.
    ///
    /// IMPORTANT - polarity and orientation:
    ///
    ///   byte 1 == background == the cell fill background color.
    ///   byte 0 == foreground == the pattern color.
    ///
    /// So it is the 0 cells that get painted with the foreground/pattern color,
    /// matching the convention used by the ECMA ST_Shd reference data.
    ///
    /// Row order matches the source bitmap exactly: row 0 is the TOP row. PDF
    /// content streams have the origin at the bottom-left with y increasing
    /// upward, so any comparison against rendered PDF output (or generation of a
    /// PDF content stream) must mirror in y (row r maps to PDF y = 7 - r). That
    /// y-flip is intentionally NOT applied here - the data is kept in bitmap
    /// orientation and the flip is the responsibility of the render/diff step.
    /// </summary>
    internal static class ExcelPatternMaskData
    {
        private static readonly Dictionary<ExcelPatternMask, byte[,]> _masks = BuildMasks();

        /// <summary>
        /// Gets the 8x8 reference mask for the given pattern.
        /// Indexed as [row, column] with row 0 = top, matching the source bitmap.
        /// </summary>
        /// <param name="pattern">The pattern to look up.</param>
        /// <returns>An 8x8 matrix where 1 = background and 0 = foreground.</returns>
        public static byte[,] GetMask(ExcelPatternMask pattern)
        {
            return _masks[pattern];
        }

        private static Dictionary<ExcelPatternMask, byte[,]> BuildMasks()
        {
            var masks = new Dictionary<ExcelPatternMask, byte[,]>();

            // DarkGray (75 gray)
            masks.Add(ExcelPatternMask.DarkGray, new byte[,]
            {
                { 0, 1, 0, 1, 0, 1, 0, 1 },
                { 1, 0, 1, 0, 1, 0, 1, 0 },
                { 0, 1, 0, 1, 0, 1, 0, 1 },
                { 1, 0, 1, 0, 1, 0, 1, 0 },
                { 0, 1, 0, 1, 0, 1, 0, 1 },
                { 1, 0, 1, 0, 1, 0, 1, 0 },
                { 0, 1, 0, 1, 0, 1, 0, 1 },
                { 1, 0, 1, 0, 1, 0, 1, 0 },
            });

            // MediumGray (50 gray)
            masks.Add(ExcelPatternMask.MediumGray, new byte[,]
            {
                { 0, 1, 1, 1, 0, 1, 1, 1 },
                { 1, 1, 0, 1, 1, 1, 0, 1 },
                { 0, 1, 1, 1, 0, 1, 1, 1 },
                { 1, 1, 0, 1, 1, 1, 0, 1 },
                { 0, 1, 1, 1, 0, 1, 1, 1 },
                { 1, 1, 0, 1, 1, 1, 0, 1 },
                { 0, 1, 1, 1, 0, 1, 1, 1 },
                { 1, 1, 0, 1, 1, 1, 0, 1 },
            });

            // LightGray (25 gray)
            masks.Add(ExcelPatternMask.LightGray, new byte[,]
            {
                { 0, 1, 1, 1, 0, 1, 1, 1 },
                { 1, 1, 1, 1, 1, 1, 1, 1 },
                { 1, 1, 0, 1, 1, 1, 0, 1 },
                { 1, 1, 1, 1, 1, 1, 1, 1 },
                { 0, 1, 1, 1, 0, 1, 1, 1 },
                { 1, 1, 1, 1, 1, 1, 1, 1 },
                { 1, 1, 0, 1, 1, 1, 0, 1 },
                { 1, 1, 1, 1, 1, 1, 1, 1 },
            });

            // Gray125 (12,5 gray)
            masks.Add(ExcelPatternMask.Gray125, new byte[,]
            {
                { 0, 1, 1, 1, 1, 1, 1, 1 },
                { 1, 1, 1, 1, 1, 1, 1, 1 },
                { 1, 1, 1, 1, 0, 1, 1, 1 },
                { 1, 1, 1, 1, 1, 1, 1, 1 },
                { 0, 1, 1, 1, 1, 1, 1, 1 },
                { 1, 1, 1, 1, 1, 1, 1, 1 },
                { 1, 1, 1, 1, 0, 1, 1, 1 },
                { 1, 1, 1, 1, 1, 1, 1, 1 },
            });

            // Gray0625 (6,25 gray)
            masks.Add(ExcelPatternMask.Gray0625, new byte[,]
            {
                { 0, 1, 1, 1, 1, 1, 1, 1 },
                { 1, 1, 1, 1, 1, 1, 1, 1 },
                { 1, 1, 1, 1, 1, 1, 1, 1 },
                { 1, 1, 1, 1, 1, 1, 1, 1 },
                { 1, 1, 1, 1, 0, 1, 1, 1 },
                { 1, 1, 1, 1, 1, 1, 1, 1 },
                { 1, 1, 1, 1, 1, 1, 1, 1 },
                { 1, 1, 1, 1, 1, 1, 1, 1 },
            });

            // DarkHorizontal (Horizontal stripe)
            masks.Add(ExcelPatternMask.DarkHorizontal, new byte[,]
            {
                { 0, 0, 0, 0, 0, 0, 0, 0 },
                { 0, 0, 0, 0, 0, 0, 0, 0 },
                { 0, 0, 0, 0, 0, 0, 0, 0 },
                { 0, 0, 0, 0, 0, 0, 0, 0 },
                { 1, 1, 1, 1, 1, 1, 1, 1 },
                { 1, 1, 1, 1, 1, 1, 1, 1 },
                { 1, 1, 1, 1, 1, 1, 1, 1 },
                { 1, 1, 1, 1, 1, 1, 1, 1 },
            });

            // DarkVertical (Vertical stripe)
            masks.Add(ExcelPatternMask.DarkVertical, new byte[,]
            {
                { 0, 0, 0, 0, 1, 1, 1, 1 },
                { 0, 0, 0, 0, 1, 1, 1, 1 },
                { 0, 0, 0, 0, 1, 1, 1, 1 },
                { 0, 0, 0, 0, 1, 1, 1, 1 },
                { 0, 0, 0, 0, 1, 1, 1, 1 },
                { 0, 0, 0, 0, 1, 1, 1, 1 },
                { 0, 0, 0, 0, 1, 1, 1, 1 },
                { 0, 0, 0, 0, 1, 1, 1, 1 },
            });

            // DarkDown (Reverse diagonal stripe)
            masks.Add(ExcelPatternMask.DarkDown, new byte[,]
            {
                { 0, 0, 0, 0, 1, 1, 1, 1 },
                { 1, 0, 0, 0, 0, 1, 1, 1 },
                { 1, 1, 0, 0, 0, 0, 1, 1 },
                { 1, 1, 1, 0, 0, 0, 0, 1 },
                { 1, 1, 1, 1, 0, 0, 0, 0 },
                { 0, 1, 1, 1, 1, 0, 0, 0 },
                { 0, 0, 1, 1, 1, 1, 0, 0 },
                { 0, 0, 0, 1, 1, 1, 1, 0 },
            });

            // DarkUp (Diagonal stripe)
            masks.Add(ExcelPatternMask.DarkUp, new byte[,]
            {
                { 1, 1, 0, 0, 0, 0, 1, 1 },
                { 1, 0, 0, 0, 0, 1, 1, 1 },
                { 0, 0, 0, 0, 1, 1, 1, 1 },
                { 0, 0, 0, 1, 1, 1, 1, 0 },
                { 0, 0, 1, 1, 1, 1, 0, 0 },
                { 0, 1, 1, 1, 1, 0, 0, 0 },
                { 1, 1, 1, 1, 0, 0, 0, 0 },
                { 1, 1, 1, 0, 0, 0, 0, 1 },
            });

            // DarkGrid (Diagonal crosshatch)
            masks.Add(ExcelPatternMask.DarkGrid, new byte[,]
            {
                { 0, 0, 0, 0, 1, 1, 1, 1 },
                { 0, 0, 0, 0, 1, 1, 1, 1 },
                { 0, 0, 0, 0, 1, 1, 1, 1 },
                { 0, 0, 0, 0, 1, 1, 1, 1 },
                { 1, 1, 1, 1, 0, 0, 0, 0 },
                { 1, 1, 1, 1, 0, 0, 0, 0 },
                { 1, 1, 1, 1, 0, 0, 0, 0 },
                { 1, 1, 1, 1, 0, 0, 0, 0 },
            });

            // DarkTrellis (Thick diagonal crosshatch)
            masks.Add(ExcelPatternMask.DarkTrellis, new byte[,]
            {
                { 0, 0, 0, 0, 0, 0, 0, 0 },
                { 1, 0, 0, 0, 0, 0, 0, 1 },
                { 1, 1, 0, 0, 0, 0, 1, 1 },
                { 1, 0, 0, 0, 0, 0, 0, 1 },
                { 0, 0, 0, 0, 0, 0, 0, 0 },
                { 0, 0, 0, 1, 1, 0, 0, 0 },
                { 0, 0, 1, 1, 1, 1, 0, 0 },
                { 0, 0, 0, 1, 1, 0, 0, 0 },
            });

            // LightHorizontal (Thin horizontal stripe)
            masks.Add(ExcelPatternMask.LightHorizontal, new byte[,]
            {
                { 0, 0, 0, 0, 0, 0, 0, 0 },
                { 0, 0, 0, 0, 0, 0, 0, 0 },
                { 1, 1, 1, 1, 1, 1, 1, 1 },
                { 1, 1, 1, 1, 1, 1, 1, 1 },
                { 1, 1, 1, 1, 1, 1, 1, 1 },
                { 1, 1, 1, 1, 1, 1, 1, 1 },
                { 1, 1, 1, 1, 1, 1, 1, 1 },
                { 1, 1, 1, 1, 1, 1, 1, 1 },
            });

            // LightVertical (Thin vertical stripe)
            masks.Add(ExcelPatternMask.LightVertical, new byte[,]
            {
                { 0, 0, 1, 1, 1, 1, 1, 1 },
                { 0, 0, 1, 1, 1, 1, 1, 1 },
                { 0, 0, 1, 1, 1, 1, 1, 1 },
                { 0, 0, 1, 1, 1, 1, 1, 1 },
                { 0, 0, 1, 1, 1, 1, 1, 1 },
                { 0, 0, 1, 1, 1, 1, 1, 1 },
                { 0, 0, 1, 1, 1, 1, 1, 1 },
                { 0, 0, 1, 1, 1, 1, 1, 1 },
            });

            // LightDown (Thin reverse diagonal stripe)
            masks.Add(ExcelPatternMask.LightDown, new byte[,]
            {
                { 0, 0, 1, 1, 1, 1, 1, 1 },
                { 1, 0, 0, 1, 1, 1, 1, 1 },
                { 1, 1, 0, 0, 1, 1, 1, 1 },
                { 1, 1, 1, 0, 0, 1, 1, 1 },
                { 1, 1, 1, 1, 0, 0, 1, 1 },
                { 1, 1, 1, 1, 1, 0, 0, 1 },
                { 1, 1, 1, 1, 1, 1, 0, 0 },
                { 0, 1, 1, 1, 1, 1, 1, 0 },
            });

            // LightUp (Thin diagonal stripe)
            masks.Add(ExcelPatternMask.LightUp, new byte[,]
            {
                { 1, 1, 1, 0, 0, 1, 1, 1 },
                { 1, 1, 0, 0, 1, 1, 1, 1 },
                { 1, 0, 0, 1, 1, 1, 1, 1 },
                { 0, 0, 1, 1, 1, 1, 1, 1 },
                { 0, 1, 1, 1, 1, 1, 1, 0 },
                { 1, 1, 1, 1, 1, 1, 0, 0 },
                { 1, 1, 1, 1, 1, 0, 0, 1 },
                { 1, 1, 1, 1, 0, 0, 1, 1 },
            });

            // LightGrid (Thin horizontal crosshatch)
            masks.Add(ExcelPatternMask.LightGrid, new byte[,]
            {
                { 0, 0, 0, 0, 0, 0, 0, 0 },
                { 0, 0, 0, 0, 0, 0, 0, 0 },
                { 0, 0, 1, 1, 1, 1, 1, 1 },
                { 0, 0, 1, 1, 1, 1, 1, 1 },
                { 0, 0, 1, 1, 1, 1, 1, 1 },
                { 0, 0, 1, 1, 1, 1, 1, 1 },
                { 0, 0, 1, 1, 1, 1, 1, 1 },
                { 0, 0, 1, 1, 1, 1, 1, 1 },
            });

            // LightTrellis (Thin diagonal crosshatch)
            masks.Add(ExcelPatternMask.LightTrellis, new byte[,]
            {
                { 0, 0, 1, 0, 0, 1, 1, 1 },
                { 1, 0, 0, 0, 1, 1, 1, 1 },
                { 1, 0, 0, 0, 1, 1, 1, 1 },
                { 0, 0, 1, 0, 0, 1, 1, 1 },
                { 0, 1, 1, 1, 0, 0, 1, 0 },
                { 1, 1, 1, 1, 1, 0, 0, 0 },
                { 1, 1, 1, 1, 1, 0, 0, 0 },
                { 0, 1, 1, 1, 0, 0, 1, 0 },
            });
            return masks;
        }
    }
}