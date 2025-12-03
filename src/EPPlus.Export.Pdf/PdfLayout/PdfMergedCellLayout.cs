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
using EPPlus.Graphics;
using EPPlus.Graphics.Math;
using System.Drawing;
using EPPlus.Export.Pdf.PdfResources;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace EPPlus.Export.Pdf.PdfLayout
{
    internal class PdfMergedCellLayout : PdfCellLayout
    {
        public PdfMergedCellLayout(PdfDictionaries dictionaries, ExcelRangeBase cell, double x, double y, double width, double height, double scaleX = 1, double scaleY = 1, double rotation = 0, Transform parent = null)
            : base(dictionaries, cell, x, y, width, height, scaleX, scaleY, rotation, parent)
        {
            this.cell = cell;
            var fill = cell.Style.Fill;
            if(!fill.HasGradient && fill.PatternType == ExcelFillStyle.None)
            {
                CellFillData.BackgroundColor = Color.White;
                CellFillData.PattenStyle = ExcelFillStyle.Solid;
                CellFillData.enhanceGridLine = true;
            }

        }

        public new void AdjustForGridLines()
        {
            if (CellFillData.BackgroundColor.Equals(Color.White) && CellFillData.PattenStyle == ExcelFillStyle.Solid)
            {
                Size = new Vector2(Size.X - GridLine.Width, Size.Y - GridLine.Width);
                LocalPosition = new Vector2(LocalPosition.X + GridLine.HalfWidth, LocalPosition.Y + GridLine.HalfWidth);
            }
        }
    }
}
