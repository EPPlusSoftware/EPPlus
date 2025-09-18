using OfficeOpenXml.PDF.Pdfhelpers;
using OfficeOpenXml.PDF.PdfResources;
using OfficeOpenXml.PDF.PdfSettings;
using OfficeOpenXml.Style;
using System.Collections.Generic;
using System.Linq;

namespace OfficeOpenXml.PDF.PdfLayout
{
    internal class PdfWorksheetLayout : PdfTransform
    {
        public PdfWorksheetLayout(ExcelWorksheet worksheet, PdfPageSettings pageSettings, Dictionary<string, PdfFontResource> fontResources)
        {
            double x = 0d, y = 0d, totalWidth = 0d;
            List<string> checkedMergedCells = new List<string>();
            for(int row = 1; row<= worksheet.Dimension._toRow; row++)
            {
                if(worksheet.Row(row).Hidden) continue;
                var height = PdfUnits.ExcelRowHeightToPoints(worksheet.Row(row).Height);
                x = 0d;
                for (int col = 1; col <= worksheet.Dimension._toCol; col++)
                {
                    if(worksheet.Column(col).Hidden) continue;
                    var width = PdfUnits.ExcelColumnWidthToPoints(worksheet.Column(col).Width);
                    var cell = worksheet.Cells[row, col];
                    if (cell.Merge)
                    {
                        HandleMergedCell(worksheet, pageSettings, fontResources, cell, checkedMergedCells, x, y);
                    }
                    else
                    {
                        HandleCell(pageSettings, fontResources, cell, x, y, width, height);
                    }
                    HandleBorders(worksheet, cell, x, y, width, height);
                    x += width;
                    totalWidth = System.Math.Max( x, totalWidth);
                }
                y += height;
            }
            HandleDrawings(worksheet);
            this.Size = new PDF.Math.Vector2(totalWidth, y);
        }

        //Create cell.
        private void HandleCell(PdfPageSettings pageSettings, Dictionary<string, PdfFontResource> fontResources, ExcelRangeBase cell, double x, double y, double width, double height)
        {
            //We add empty cells for gridline calculation later. We just marked them for deletion by addng * to their name.
            string deleteMark = !cell.IsEmpty() || cell.Worksheet.ExistsStyleInner(cell._fromRow, cell._toCol) ? "" : "*";
            var cl0 = new PdfCellLayout(cell, x, y, width, height, 1, 1, 0, this);
            cl0.Name = cell.Address + deleteMark;
            cl0.Z = 1;
            AddCellContent(pageSettings, fontResources, cell, x, y, width, height, 2);
        }

        //Create merged cell.
        private void HandleMergedCell(ExcelWorksheet worksheet, PdfPageSettings pageSettings, Dictionary<string, PdfFontResource> fontResources, ExcelRangeBase cell, List<string> checkedMergedCells, double x, double y)
        {
            string mergeAddress = worksheet.MergedCells[cell.Start.Row, cell.Start.Column];
            if (!checkedMergedCells.Contains(mergeAddress))
            {
                double height = 0, width = 0;
                ExcelAddressBase address = new ExcelAddressBase(mergeAddress);
                for (int k = address._fromRow; k <= address._toRow; k++)
                {
                    height += PdfUnits.ExcelRowHeightToPoints(worksheet.Row(k).Height);
                }
                for (int l = address._fromCol; l <= address._toCol; l++)
                {
                    width += PdfUnits.ExcelColumnWidthToPoints(worksheet.Column(l).Width);
                }
                var mergedCell = AddChild(new PdfMergedCellLayout(worksheet.Cells[address._fromRow, address._fromCol], x, y, width, height));
                mergedCell.Name = cell.Address;
                mergedCell.Z = 5;
                AddCellContent(pageSettings, fontResources, cell, x, y, width, height, 6);
                checkedMergedCells.Add(mergeAddress);
            }
        }

        //Create content.
        private void AddCellContent(PdfPageSettings pageSettings, Dictionary<string, PdfFontResource> fontResources, ExcelRangeBase cell, double x, double y, double width, double height, int zOrder)
        {
            if (!string.IsNullOrEmpty(cell.Text))
            {
                var cellContent = new PdfCellContentLayout(cell, pageSettings, x, y, width, height, 1, 1, 0, this, fontResources);
                cellContent.Name = cell.Address;
                cellContent.Z = zOrder;
            }
        }

        //Create borders.
        private void HandleBorders(ExcelWorksheet worksheet, ExcelRangeBase cell, double x, double y, double width, double height)
        {
            bool allNone = new[] { cell.Style.Border.Top.Style, cell.Style.Border.Bottom.Style, cell.Style.Border.Left.Style, cell.Style.Border.Right.Style, cell.Style.Border.Diagonal.Style }.All(s => s == ExcelBorderStyle.None);
            if (!allNone)
            {
                var clb0 = new PdfCellBorderLayout(cell, worksheet.Dimension, x, y, width, height, 1, 1, 0, this);
                clb0.Name = cell.Address;
                clb0.Z = 7;
            }
        }

        //Create drawings.
        private void HandleDrawings(ExcelWorksheet worksheet)
        {
            foreach (var drawing in worksheet.Drawings) //NOT IMPLEMENTED
            {
                var drawLayout = AddChild(new PdfDrawingLayout(drawing, drawing.Position.X, drawing.Position.Y, drawing._width, drawing._height));
                drawLayout.Z = 10;
                drawLayout.Name = "Drawing " + drawing.Name;
            }
        }
    }
}
