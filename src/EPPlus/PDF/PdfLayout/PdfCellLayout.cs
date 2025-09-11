using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using OfficeOpenXml.Interfaces.Drawing.Text;
using OfficeOpenXml.PDF.Math;
using OfficeOpenXml.PDF.PdfGraphics;
using OfficeOpenXml.PDF.Pdfhelpers;
using OfficeOpenXml.PDF.PdfSettings;
using OfficeOpenXml.Style;

namespace OfficeOpenXml.PDF.PdfLayout
{
    internal class PdfCellLayout : PdfTransform
    {
        public PdfCellFillData CellFillData;

        public PdfCellLayout(ExcelRangeBase cell, double x, double y, double width, double height, double scaleX = 1, double scaleY = 1, double rotation = 0, PdfTransform parent = null)
            :base(x, y, width, height, scaleX, scaleY, rotation, parent )
        {
            if (cell != null)
            {
                CellFillData = new PdfCellFillData();
                var bkgc = cell.Style.Fill.BackgroundColor;
                if (string.IsNullOrEmpty(bkgc.LookupColor()) && !string.IsNullOrEmpty(cell.Text))
                {
                    CellFillData.BackgroundColor = PdfColor.None;
                }
                else
                {
                    CellFillData.BackgroundColor = new PdfColor(bkgc.LookupColor());
                }
                CellFillData.PattenStyle = cell.Style.Fill.PatternType;
                CellFillData.PatternColor = new PdfColor(cell.Style.Fill.PatternColor.LookupColor());
            }
        }

        //Adjust size and position slightly for aesthetics.
        public void AdjustForGridLines()
        {
            Size = new Vector2(Size.X + GridLine.HalfWidth, Size.Y + GridLine.HalfWidth);
            LocalPosition = new Vector2(LocalPosition.X + GridLine.FourthWidth, LocalPosition.Y + GridLine.HalfWidth + GridLine.FourthWidth);
        }
    }
}
