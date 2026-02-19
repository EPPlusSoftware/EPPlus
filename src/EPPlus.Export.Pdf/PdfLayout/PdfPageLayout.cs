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
using OfficeOpenXml;
using System.Collections.Generic;
using System.Linq;

namespace EPPlus.Export.Pdf.PdfLayout
{
    internal class PdfPageLayout : Transform
    {
        internal List<GridLine> GridLines = new List<GridLine>();
        internal List<GridLine> BorderLines = new List<GridLine>();

        public PdfPageLayout(double x, double y, double width, double height)
            :base(x, y, width, height) { }

        internal void GenerateGridLines()
        {
            HashSet<double> xCoords = new HashSet<double>();
            HashSet<double> yCoords = new HashSet<double>();
            // Collect all unique X and Y coordinates.
            var cells = ChildObjects.Where(x => x is PdfCellLayout).ToList();
            foreach (var c in cells)
            {
                if (c is PdfMergedCellLayout) continue;
                yCoords.Add(c.LocalPosition.Y + c.Size.Y);  // Top
                yCoords.Add(c.LocalPosition.Y);             // Bottom
                xCoords.Add(c.LocalPosition.X);             // Left
                xCoords.Add(c.LocalPosition.X + c.Size.X);  // Right
            }
            double minX = cells.Where(c => c is not PdfMergedCellLayout).Min(c => c.LocalPosition.X);
            double maxX = cells.Where(c => c is not PdfMergedCellLayout).Max(c => c.LocalPosition.X + c.Size.X);
            double minY = cells.Where(c => c is not PdfMergedCellLayout).Min(c => c.LocalPosition.Y);
            double maxY = cells.Where(c => c is not PdfMergedCellLayout).Max(c => c.LocalPosition.Y + c.Size.Y);
            foreach (var x in xCoords.OrderBy(v => v))
            {
                var line = new GridLine(x, minY, x, maxY);
                if (x == minX || x == maxX)
                    BorderLines.Add(line);
                else
                    GridLines.Add(line);
            }
            foreach (var y in yCoords.OrderBy(v => v))
            {
                var line = new GridLine(minX, y, maxX, y);
                if (y == minY || y == maxY)
                    BorderLines.Add(line);
                else
                    GridLines.Add(line);
            }
        }


        internal void GenerateVerticalGridLines(ExcelWorksheet ws)
        {
            int addedColumns = PdfWorksheetLayout.AddColumnsForNonWrappedText(ws);
            Dictionary<string, List<double>> vertGridLines = new Dictionary<string, List<double>>();
            for (int col = 1; col <= ws.Dimension._toCol + addedColumns; col++)
            {
                List<double> lengths = new List<double>();
                double length = 0;
                string name = "";
                for (int row = 1; row <= ws.Dimension._toRow; row++)
                {
                    var cell = ws.Cells[row, col];
                    name = cell.Address;
                    var layouts = ChildObjects.Where(x => x.Name == cell.Address || x.Name == cell.Address +"_c" || x.Name == cell.Address + "_m" || x.Name == cell.Address + "*").ToList();
                    PdfMergedCellLayout m = null;
                    PdfCellContentLayout c = null;
                    PdfCellLayout l = null;
                    foreach (var layout in layouts)
                    {
                        if (layout is PdfMergedCellLayout) m = (PdfMergedCellLayout)layout;
                        else if (layout is PdfCellContentLayout) c = (PdfCellContentLayout)layout;
                        else if (layout is PdfCellLayout) l = (PdfCellLayout)layout;
                    }
                    if (l != null)
                    {
                        length += l.Size.Y;
                    }
                    else if (m != null)
                    {
                        //ExcelAddressBase address = new ExcelAddressBase(ws.MergedCells[cell.Start.Row, cell.Start.Column]);
                        //if (address._fromCol == col)
                        //{
                            length += m.Size.Y;
                        //}
                        //else
                        //{
                            //length = 0;
                        //}
                    }
                    else if (l == null && m == null)
                    {
                        if (length > 0)
                        {
                            lengths.Add(length);
                            length = 0;
                        }
                    }
                }
                lengths.Add(length);
                vertGridLines.Add(name, lengths); //error in B column where merged cell starts
            }
        }
    }
}
