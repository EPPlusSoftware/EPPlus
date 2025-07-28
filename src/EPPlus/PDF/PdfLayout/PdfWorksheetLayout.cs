using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup;
using OfficeOpenXml.PDF.Pdfhelpers;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.PDF.PdfLayout
{
    internal class PdfWorksheetLayout : PdfTransform
    {
        internal ExcelWorksheet ws;


        public PdfWorksheetLayout(ExcelWorksheet worksheet)
        {
            this.ws = worksheet;
            double x = 0d;
            double y = 0d;
            double totalWidth = 0;
            List<string> checkedMergedCells = new List<string>();
            for(int i = 1; i<= ws.Dimension._toRow; i++)
            {
                var height = ws.Row(i).Height;
                for (int j = 1; j <= ws.Dimension._toCol; j++)
                {
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
                                mcWidth += (ws.Column(l).Width);
                            }
                            var cl1 = AddChild(new PdfCellLayout(ws.Cells[address._fromRow, address._fromCol].Value, x, y, mcWidth, mcHeight));
                            cl1.Z = 1;
                            cl1.Name = cell.Address;
                            checkedMergedCells.Add(ws.MergedCells[i, j]);
                        }
                    }
                    var width = PdfUnits.ExcelColumnWidthToPoints(ws.Column(j).Width);
                    var cl0 =  AddChild(new PdfCellLayout((isMerged ? null : cell.Value), x, y, width, height));
                    cl0.Z = 1;
                    cl0.Name = cell.Address;
                    x+= width;
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
                drawLayout.Name = drawing.Name;
            }
        }
    }
}
