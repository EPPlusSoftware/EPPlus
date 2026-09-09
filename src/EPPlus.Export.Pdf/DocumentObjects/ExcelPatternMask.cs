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