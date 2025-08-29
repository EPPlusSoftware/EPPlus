using OfficeOpenXml.PDF.PdfGraphics;
using OfficeOpenXml.PDF.PdfSettings;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.PDF.PdfLayout
{
    internal class PdfCellBorderLayout : PdfTransform
    {
        public PdfCellBordersData BorderData;

        PdfPageSettings PageSettings;

        public PdfCellBorderLayout() { }

        public PdfCellBorderLayout(ExcelRangeBase cell, ExcelRangeBase dimension, PdfPageSettings pageSettings, double x, double y, double width, double height, double scaleX = 1, double scaleY = 1, double rotation = 0, PdfTransform parent = null)
            : base(x, y, width, height, scaleX, scaleY, rotation, parent)
        {
            if (cell != null)
            {
                BorderData = new PdfCellBordersData();
                BorderData.Top.BorderStyle = cell.Style.Border.Top.Style;
                BorderData.Top.BorderColor = new PdfColor(cell.Style.Border.Top.Color.LookupColor(cell.Style.Border));
                BorderData.Bottom.BorderStyle = cell.Style.Border.Bottom.Style;
                BorderData.Bottom.BorderColor = new PdfColor(cell.Style.Border.Bottom.Color.LookupColor(cell.Style.Border));
                BorderData.Left.BorderStyle = cell.Style.Border.Left.Style;
                BorderData.Left.BorderColor = new PdfColor(cell.Style.Border.Left.Color.LookupColor(cell.Style.Border));
                BorderData.Right.BorderStyle = cell.Style.Border.Right.Style;
                BorderData.Right.BorderColor = new PdfColor(cell.Style.Border.Right.Color.LookupColor(cell.Style.Border));
                BorderData.DiagonalUp.BorderStyle = cell.Style.Border.DiagonalUp ? cell.Style.Border.Diagonal.Style : ExcelBorderStyle.None;
                BorderData.DiagonalUp.BorderColor = new PdfColor(cell.Style.Border.Diagonal.Color.LookupColor(cell.Style.Border));
                BorderData.DiagonalDown.BorderStyle = cell.Style.Border.DiagonalDown ? cell.Style.Border.Diagonal.Style : ExcelBorderStyle.None;
                BorderData.DiagonalDown.BorderColor = new PdfColor(cell.Style.Border.Diagonal.Color.LookupColor(cell.Style.Border));
            }
            PageSettings = pageSettings;
            if(pageSettings.ShowGridLines)
            {
                if (BorderData.Top.BorderStyle == ExcelBorderStyle.None)
                {
                    BorderData.Top.BorderStyle = cell._fromRow == dimension._fromRow ? ExcelBorderStyle.GridOuter : ExcelBorderStyle.None;
                    BorderData.Top.BorderColor = PdfColor.Black;
                }
                if (BorderData.Bottom.BorderStyle == ExcelBorderStyle.None)
                {
                    BorderData.Bottom.BorderStyle = cell._toRow == dimension._toRow ? ExcelBorderStyle.GridOuter : ExcelBorderStyle.GridInner;
                    BorderData.Bottom.BorderColor =/* BorderData.Bottom.BorderStyle == ExcelBorderStyle.GridInner ? PdfColor.Gray :*/ PdfColor.Black;
                }
                if (BorderData.Left.BorderStyle == ExcelBorderStyle.None)
                {
                    BorderData.Left.BorderStyle = cell._fromCol == dimension._fromCol  ? ExcelBorderStyle.GridOuter : ExcelBorderStyle.None;
                    BorderData.Left.BorderColor = PdfColor.Black;
                }
                if (BorderData.Right.BorderStyle == ExcelBorderStyle.None)
                {
                    BorderData.Right.BorderStyle = cell._toCol == dimension._toCol ? ExcelBorderStyle.GridOuter : ExcelBorderStyle.GridInner;
                    BorderData.Right.BorderColor = /*BorderData.Right.BorderStyle == ExcelBorderStyle.GridInner ? PdfColor.Gray :*/ PdfColor.Black;
                }
            }
        }
    }
}

