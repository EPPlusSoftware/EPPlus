using OfficeOpenXml.PDF.Math;
using OfficeOpenXml.PDF.PdfGraphics;
using OfficeOpenXml.PDF.PdfSettings;
using OfficeOpenXml.Style;

namespace OfficeOpenXml.PDF.PdfLayout
{
    internal class PdfCellBorderLayout : PdfTransform, ILayout
    {
        public PdfCellBordersData BorderData;

        public PdfCellBorderLayout() { }

        public PdfCellBorderLayout(ExcelRangeBase cell, ExcelRangeBase dimension, double x, double y, double width, double height, double scaleX = 1, double scaleY = 1, double rotation = 0, PdfTransform parent = null)
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
        }

        public void ConvertCoordinates(PdfPageSettings pageSettings)
        {
            LocalPosition = new Vector2(LocalPosition.X, (pageSettings.PageSize.HeightPu - System.Math.Abs(LocalPosition.Y)) - Size.Y);
        }
    }
}

