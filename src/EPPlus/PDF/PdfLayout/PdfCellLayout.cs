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
        public PdfCellBordersData BorderData;

        PdfPageSettings PageSettings;

        public PdfCellLayout() { }

        public PdfCellLayout(ExcelRangeBase cell, PdfPageSettings pageSettings, double x, double y, double width, double height, double scaleX = 1, double scaleY = 1, double rotation = 0, PdfTransform parent = null)
            :base(x, y, width, height, scaleX, scaleY, rotation, parent )
        {
            if (cell != null)
            {
                CellFillData = new PdfCellFillData();
                var bkgc = cell.Style.Fill.BackgroundColor;
                if (string.IsNullOrEmpty(bkgc.LookupColor()) && !string.IsNullOrEmpty(cell.Text))
                {
                    CellFillData.BackgroundColor = PdfColor.White;
                }
                else
                {
                    CellFillData.BackgroundColor = new PdfColor(bkgc.LookupColor());
                }
                CellFillData.PattenStyle = cell.Style.Fill.PatternType;
                CellFillData.PatternColor = new PdfColor(cell.Style.Fill.PatternColor.LookupColor());
                BorderData = new PdfCellBordersData();
                BorderData.Top.BorderStyle = cell.Style.Border.Top.Style;
                BorderData.Top.BorderColor = new PdfColor(cell.Style.Border.Top.Color.LookupColor());
                BorderData.Bottom.BorderStyle = cell.Style.Border.Bottom.Style;
                BorderData.Bottom.BorderColor = new PdfColor(cell.Style.Border.Bottom.Color.LookupColor());
                BorderData.Left.BorderStyle = cell.Style.Border.Left.Style;
                BorderData.Left.BorderColor = new PdfColor(cell.Style.Border.Left.Color.LookupColor());
                BorderData.Right.BorderStyle = cell.Style.Border.Right.Style;
                BorderData.Right.BorderColor = new PdfColor(cell.Style.Border.Right.Color.LookupColor());
                BorderData.DiagonalUp.BorderStyle = cell.Style.Border.DiagonalUp ? cell.Style.Border.Diagonal.Style : ExcelBorderStyle.None;
                BorderData.DiagonalUp.BorderColor = new PdfColor(cell.Style.Border.Diagonal.Color.LookupColor());
                BorderData.DiagonalDown.BorderStyle = cell.Style.Border.DiagonalDown ? cell.Style.Border.Diagonal.Style : ExcelBorderStyle.None;
                BorderData.DiagonalDown.BorderColor = new PdfColor(cell.Style.Border.Diagonal.Color.LookupColor());
            }
            PageSettings = pageSettings;
        }
    }
}
