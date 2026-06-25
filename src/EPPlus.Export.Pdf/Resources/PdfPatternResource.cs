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
using EPPlus.Export.Pdf.Layout;
using EPPlus.Export.Pdf.DocumentObjects.Patterns;
using EPPlus.Export.Pdf.DocumentObjects;
using EPPlus.Export.Pdf.Enums;

namespace EPPlus.Export.Pdf.Resources
{
    internal class PdfPatternResource : PdfResource
    {
        internal int objectNumber;
        internal PdfCellFillData CellFillData;

        // All mask-based patterns use an 8x8 tile in pattern space. The mask is the
        // single source of geometry, so BBox and the tiling step are 8x8 for every
        // pattern.
        private static readonly double[] PatternBBox = new double[] { 0d, 0d, 8d, 8d };
        private const double PatternStepX = 8d;
        private const double PatternStepY = 8d;

        // The 8x8 pattern space is scaled down to match the physical tile size
        // Excel uses (0.75 pt for a full tile, i.e. 0.75 / 8 per pattern unit).
        // The scale is positive on both axes: the renderer already mirrors the
        // mask in y, so no y-flip is applied here (that would double-flip).
        private const double PatternScale = 0.75d / 8d;
        private static readonly double[] PatternMatrix =
            new double[] { PatternScale, 0d, 0d, PatternScale, 0d, 0d };

        public PdfPatternResource(int labelNumber, PdfCellFillData cellFillData)
            : base("P", labelNumber)
        {
            CellFillData = cellFillData;
        }

        public PdfPattern GetPatternObject(int objectNumber, int version = 0)
        {
            this.objectNumber = objectNumber;
                        // None and Solid are not patterns; they are handled as special cases
            // elsewhere in the export.
            if (CellFillData.PatternStyle == ExcelFillStyle.None ||
                CellFillData.PatternStyle == ExcelFillStyle.Solid)
            {
                return null;
            }
            ExcelPatternMask mask;
            if (!TryGetMask(CellFillData.PatternStyle, out mask))
            {
                return null;
            }
            var fill = new PdfPatternMaskFill(mask, CellFillData.PatternColor, CellFillData.BackgroundColor);
            var pattern = new PdfTilingPattern(objectNumber, fill, PatternBBox, PatternStepX, PatternStepY, version);
            pattern.Matrix = PatternMatrix;
            return pattern;
        }

        /// <summary>
        /// Maps an Excel cell fill style to the corresponding 8x8 pattern mask.
        /// Returns false for styles that have no mask (e.g. None/Solid, or any
        /// style not rendered as a tiling pattern).
        /// </summary>
        private static bool TryGetMask(ExcelFillStyle style, out ExcelPatternMask mask)
        {
            switch (style)
            {
                case ExcelFillStyle.DarkGray: mask = ExcelPatternMask.DarkGray; return true;
                case ExcelFillStyle.MediumGray: mask = ExcelPatternMask.MediumGray; return true;
                case ExcelFillStyle.LightGray: mask = ExcelPatternMask.LightGray; return true;
                case ExcelFillStyle.Gray125: mask = ExcelPatternMask.Gray125; return true;
                case ExcelFillStyle.Gray0625: mask = ExcelPatternMask.Gray0625; return true;
                case ExcelFillStyle.DarkHorizontal: mask = ExcelPatternMask.DarkHorizontal; return true;
                case ExcelFillStyle.DarkVertical: mask = ExcelPatternMask.DarkVertical; return true;
                case ExcelFillStyle.DarkDown: mask = ExcelPatternMask.DarkDown; return true;
                case ExcelFillStyle.DarkUp: mask = ExcelPatternMask.DarkUp; return true;
                case ExcelFillStyle.DarkGrid: mask = ExcelPatternMask.DarkGrid; return true;
                case ExcelFillStyle.DarkTrellis: mask = ExcelPatternMask.DarkTrellis; return true;
                case ExcelFillStyle.LightHorizontal: mask = ExcelPatternMask.LightHorizontal; return true;
                case ExcelFillStyle.LightVertical: mask = ExcelPatternMask.LightVertical; return true;
                case ExcelFillStyle.LightDown: mask = ExcelPatternMask.LightDown; return true;
                case ExcelFillStyle.LightUp: mask = ExcelPatternMask.LightUp; return true;
                case ExcelFillStyle.LightGrid: mask = ExcelPatternMask.LightGrid; return true;
                case ExcelFillStyle.LightTrellis: mask = ExcelPatternMask.LightTrellis; return true;
                default:
                    mask = ExcelPatternMask.DarkGray;
                    return false;
            }
        }
    }
}