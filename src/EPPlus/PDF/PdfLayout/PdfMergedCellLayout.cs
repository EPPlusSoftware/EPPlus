using OfficeOpenXml.PDF.PdfGraphics;

namespace OfficeOpenXml.PDF.PdfLayout
{
    internal class PdfMergedCellLayout : PdfCellLayout
    {
        public PdfMergedCellLayout(ExcelRangeBase cell, double x, double y, double width, double height, double scaleX = 1, double scaleY = 1, double rotation = 0, PdfTransform parent = null)
            : base(cell, x, y, width, height, scaleX, scaleY, rotation, parent)
        {
            this.cell = cell;
            //Test removing this and see if adding clipping rect to merged cells work. It wont work.. find some other way.
            if (CellFillData.BackgroundColor.Equals(PdfColor.None))
            {
                CellFillData.BackgroundColor = PdfColor.White;
            }
        }
    }
}
