using OfficeOpenXml.FormulaParsing.Excel.Functions.Text;
using OfficeOpenXml.PDF.PdfSettings;
using System.Collections.Generic;
using System.Linq;

namespace OfficeOpenXml.PDF.PdfLayout
{
    internal class PdfPageLayout : PdfTransform
    {
        internal ExcelRangeBase Range;

        public PdfPageLayout(double x, double y, double width, double height)
            :base(x, y, width, height)
        {
        }

        internal void FixGridLines(PdfPageSettings pageSettings, ExcelWorksheet ws)
        {
            if (pageSettings.ShowGridLines)
            {
                //TODO: Check neighbour cell for border style and skip grid line if one exsist. Maybe do this in the WorksheetLayout class
                var topBorder = ChildObjects.Where(t => ws.Cells[t.Name]._fromRow == Range.Start.Row).OfType<PdfCellBorderLayout>();
                var bottomBorder = ChildObjects.Where(t => ws.Cells[t.Name]._toRow == Range.End.Row).OfType<PdfCellBorderLayout>();
                var leftBorder = ChildObjects.Where(t => t.Name.Contains(ExcelCellBase.GetColumnLetter(Range.Start.Column))).OfType<PdfCellBorderLayout>();
                var rightBorder = ChildObjects.Where(t => t.Name.Contains(ExcelCellBase.GetColumnLetter(Range.End.Column))).OfType<PdfCellBorderLayout>();
                foreach (PdfCellBorderLayout top in topBorder)
                {
                    if (top.BorderData.Top.BorderStyle == Style.ExcelBorderStyle.None)
                    {
                        top.BorderData.Top.BorderStyle = Style.ExcelBorderStyle.GridOuter;
                        top.BorderData.Top.BorderColor = PdfGraphics.PdfColor.Black;
                    }
                }
                foreach (PdfCellBorderLayout bottom in bottomBorder)
                {
                    if (bottom.BorderData.Bottom.BorderStyle == Style.ExcelBorderStyle.None || bottom.BorderData.Bottom.BorderStyle == Style.ExcelBorderStyle.GridInner)
                    {
                        bottom.BorderData.Bottom.BorderStyle = Style.ExcelBorderStyle.GridOuter;
                        bottom.BorderData.Bottom.BorderColor = PdfGraphics.PdfColor.Black;
                    }
                }
                foreach (PdfCellBorderLayout left in leftBorder)
                {
                    if (left.BorderData.Left.BorderStyle == Style.ExcelBorderStyle.None)
                    {
                        left.BorderData.Left.BorderStyle = Style.ExcelBorderStyle.GridOuter;
                        left.BorderData.Left.BorderColor = PdfGraphics.PdfColor.Black;
                    }
                }
                foreach (PdfCellBorderLayout right in rightBorder)
                {
                    if (right.BorderData.Right.BorderStyle == Style.ExcelBorderStyle.None || right.BorderData.Right.BorderStyle == Style.ExcelBorderStyle.GridInner)
                    {
                        right.BorderData.Right.BorderStyle = Style.ExcelBorderStyle.GridOuter;
                        right.BorderData.Right.BorderColor = PdfGraphics.PdfColor.Black;
                    }
                }
            }
        }

    }
}
