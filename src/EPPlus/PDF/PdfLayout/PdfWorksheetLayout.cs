using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup;
using OfficeOpenXml.PDF.PdfFontData;
using OfficeOpenXml.PDF.Pdfhelpers;
using OfficeOpenXml.PDF.PdfSettings;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.PDF.PdfLayout
{
    internal class PdfWorksheetLayout : PdfTransform
    {
        internal ExcelWorksheet ws;


        public PdfWorksheetLayout(ExcelWorksheet worksheet, PdfPageSettings pageSettings, PdfContentBounds bounds, Dictionary<string, PdfFontResource> fontResources)
        {
            this.ws = worksheet;
            double x = 0d;
            double y = 0d;
            double totalWidth = 0d;
            List<string> checkedMergedCells = new List<string>();
            for(int i = 1; i<= ws.Dimension._toRow; i++)
            {
                if(ws.Row(i).Hidden) { continue; }
                var height = ws.Row(i).Height;
                for (int j = 1; j <= ws.Dimension._toCol; j++)
                {
                    if(ws.Column(j).Hidden) { continue; }
                    var cell = ws.Cells[i, j];
                    bool isMerged = false;
                    if (cell.Merge)
                    {
                        isMerged = true;
                        if (!checkedMergedCells.Contains(ws.MergedCells[i, j]))
                        {
                            double mcHeight = 0;
                            double mcWidth = 0;
                            ExcelAddressBase address = new ExcelAddressBase(ws.MergedCells[i, j]);
                            for (int k = address._fromRow; k <= address._toRow; k++)
                            {
                                mcHeight += ws.Row(k).Height;
                            }
                            for (int l = address._fromCol; l <= address._toCol; l++)
                            {
                                mcWidth += PdfUnits.ExcelColumnWidthToPoints(ws.Column(l).Width);
                            }
                            var cl1 = AddChild(new PdfMergedCellLayout(ws.Cells[address._fromRow, address._fromCol], pageSettings, x, y, mcWidth, mcHeight));
                            cl1.Z = 5;
                            cl1.Name = cell.Address;
                            if (!string.IsNullOrEmpty(cell.Text))
                            {
                                var clc1 = new PdfCellContentLayout(isMerged ? null : cell, pageSettings, x, y, mcWidth, mcHeight, 1, 1, 0, this, fontResources);
                                clc1.Z = 6;
                                clc1.Name = cell.Address;
                            }
                            checkedMergedCells.Add(ws.MergedCells[i, j]);
                        }
                    }
                    var width = PdfUnits.ExcelColumnWidthToPoints(ws.Column(j).Width);
                    if (!cell.IsEmpty() || cell.Worksheet.ExistsStyleInner(cell._fromRow, cell._toCol))
                    {
                        var cl0 = new PdfCellLayout((isMerged ? null : cell), pageSettings, x, y, width, height, 1, 1, 0, this);
                        cl0.Z = 1;
                        cl0.Name = cell.Address;
                        if (!string.IsNullOrEmpty(cell.Text))
                        {
                            var clc0 = new PdfCellContentLayout(isMerged ? null : cell, pageSettings, x, y, width, height, 1, 1, 0, this, fontResources);

                            clc0.Z = 2;
                            clc0.Name = cell.Address;
                        }
                    }
                    x += width;
                    if (x > totalWidth)
                    {
                        totalWidth = x;
                    }
                }
                y += height;
                x = 0d;
            }
            this.Size = new PDF.Math.Vector2(totalWidth, y);
            foreach(var drawing in ws.Drawings)
            {
                //TODO: convert size and position to pdf coords needed.
                var drawLayout = AddChild(new PdfDrawingLayout(drawing, drawing.Position.X, drawing.Position.Y, drawing._width, drawing._height));
                drawLayout.Z = 10;
                drawLayout.Name = "Drawing " + drawing.Name;
            }
        }
    }
}
