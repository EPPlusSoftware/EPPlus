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
using System.Collections.Generic;
using System.Diagnostics;

namespace EPPlus.Export.Pdf.Layout
{
    [DebuggerDisplay("Cell: {Name}")]
    internal struct PageMap
    {
        public string Name;
        public int row;
        public int col;
        public PdfCellLayout cell;
        public PdfCellContentLayout content;
        public PdfCellBorderLayout border;
        public enum CellType
        {
            Normal,
            Merged,
        }
        public CellType Type;

        //Scoops up text spill from adjecent cells
        public double LeftTextBucketSpill;
        public double RightTextBucketSpill;
    }

    internal class VerticalLineRun
    {
        public int Col;
        public int RowStart;
        public int RowEnd;
    }

    internal class HorizontalLineRun
    {
        public int Row;
        public int ColStart;
        public int ColEnd;
    }

    [DebuggerDisplay("{Name}")]
    internal class PdfPageLayout : Transform
    {
        internal List<GridLine> GridLines = new List<GridLine>();
        internal List<GridLine> BorderLines = new List<GridLine>();
        public List<GridLine> PrintTitleGridLines = new List<GridLine>();

        public int FromRow = int.MaxValue;
        public int ToRow = 0;
        public int FromCol = int.MaxValue;
        public int ToCol = 0;
        public int RowCount = 0;
        public int ColCount = 0;

        public PageMap[,] Map;

        public double ContentTop;
        public double ContentBottom;
        public double ContentLeft;
        public double ContentRight;
        public double ContentX;
        public double ContentHeight;

        public double HeadingWidth;
        public double HeadingHeight;

        public double PrintTitleWidth;
        public double PrintTitleHeight;

        public bool isCommentsPage = false;

        public PdfPageLayout(double x, double y, double width, double height)
            : base(x, y, width, height) { }

        //public void AddCell(Transform cell)
        //{
        //    if (cell is PdfCellLayout cl)
        //    {
        //        FromRow = Math.Min(FromRow, cl.cell._fromRow);
        //        ToRow = Math.Max(ToRow, cl.cell._toRow);
        //        FromCol = Math.Min(FromCol, cl.cell._fromCol);
        //        ToCol = Math.Max(ToCol, cl.cell._toCol);
        //    }
        //    ChildObjects[0].AddChild(cell);
        //}

        //public void CreateMap()
        //{
        //    RowCount = ToRow - FromRow + 1;
        //    ColCount = ToCol - FromCol + 1;
        //    Map = new PageMap[RowCount, ColCount];
        //}

        //internal void GenerateGridLines()
        //{
        //    HashSet<double> xCoords = new HashSet<double>();
        //    HashSet<double> yCoords = new HashSet<double>();
        //    // Collect all unique X and Y coordinates.
        //    var cells = ChildObjects.Where(x => x is PdfCellLayout).ToList();
        //    foreach (var c in cells)
        //    {
        //        if (c is PdfMergedCellLayout) continue;
        //        yCoords.Add(c.LocalPosition.Y + c.Size.Y);  // Top
        //        yCoords.Add(c.LocalPosition.Y);             // Bottom
        //        xCoords.Add(c.LocalPosition.X);             // Left
        //        xCoords.Add(c.LocalPosition.X + c.Size.X);  // Right
        //    }
        //    double minX = cells.Where(c => c is not PdfMergedCellLayout).Min(c => c.LocalPosition.X);
        //    double maxX = cells.Where(c => c is not PdfMergedCellLayout).Max(c => c.LocalPosition.X + c.Size.X);
        //    double minY = cells.Where(c => c is not PdfMergedCellLayout).Min(c => c.LocalPosition.Y);
        //    double maxY = cells.Where(c => c is not PdfMergedCellLayout).Max(c => c.LocalPosition.Y + c.Size.Y);
        //    foreach (var x in xCoords.OrderBy(v => v))
        //    {
        //        var line = new GridLine(x, minY, x, maxY);
        //        if (x == minX || x == maxX)
        //            BorderLines.Add(line);
        //        else
        //            GridLines.Add(line);
        //    }
        //    foreach (var y in yCoords.OrderBy(v => v))
        //    {
        //        var line = new GridLine(minX, y, maxX, y);
        //        if (y == minY || y == maxY)
        //            BorderLines.Add(line);
        //        else
        //            GridLines.Add(line);
        //    }
        //}

        //internal void GenerateVerticalGridLines(ExcelWorksheet ws)
        //{
        //    int addedColumns = PdfWorksheetLayout.AddColumnsForNonWrappedText(ws);
        //    bool resetStart = true;
        //    var startX = 0d;
        //    var startY = 0d;
        //    var endX = 0d;
        //    var endY = 0d;
        //    for (int col = FromCol+1; col <= ToCol; col++)
        //    {
        //        double length = 0;
        //        //string name = "";
        //        //var f = ws.Cells[FromRow, col];
        //        //var start = ChildObjects.Where(x => x.Name == f.Address || x.Name == f.Address + "_m" || x.Name == f.Address + "*").ToList();
        //        for (int row = FromRow; row <= ToRow; row++)
        //        {
        //            var cell = ws.Cells[row, col];
        //            var name = cell.Address;
        //            var layouts = ChildObjects.Where(x => x.Name == cell.Address || x.Name == cell.Address + "_m" || x.Name == cell.Address + "*").ToList();
        //            PdfMergedCellLayout m = null;
        //            PdfCellLayout l = null;
        //            foreach (var layout in layouts)
        //            {
        //                if (layout is PdfMergedCellLayout) m = (PdfMergedCellLayout)layout;
        //                else if (layout is PdfCellLayout) l = (PdfCellLayout)layout;
        //            }
        //            if (l != null)
        //            {
        //                if (l.LeftTextSpillLength > 0)
        //                {
        //                    length = 0;
        //                    if (startX != 0d)
        //                    {
        //                        var line = new GridLine(startX, startY, endX, endY);
        //                        GridLines.Add(line);
        //                        resetStart = true;
        //                    }
        //                }
        //                else
        //                {
        //                    length += l.Size.Y;
        //                    endX = l.LocalPosition.X;
        //                    endY = l.LocalPosition.Y;
        //                    if (resetStart)
        //                    {
        //                        startX = l.LocalPosition.X;
        //                        startY = l.LocalPosition.Y + l.Size.Y;
        //                        resetStart = false;
        //                    }
        //                }
        //            }
        //            else if (m != null)
        //            {
        //                length += m.Size.Y;
        //                endX = m.LocalPosition.X;
        //                endY = m.LocalPosition.Y;
        //                if (resetStart)
        //                {
        //                    startX = m.LocalPosition.X;
        //                    startY = m.LocalPosition.Y + m.Size.Y;
        //                    resetStart = false;
        //                }
        //            }
        //            else if (l == null && m == null)
        //            {
        //                if (length > 0)
        //                {
        //                    length = 0;
        //                    if (startX != 0d)
        //                    {
        //                        var line = new GridLine(startX, startY, endX, endY);
        //                        GridLines.Add(line);
        //                        resetStart = true;
        //                    }
        //                }
        //            }
        //        }
        //        if (startX != 0d)
        //        {
        //            var line2 = new GridLine(startX, startY, endX, endY);
        //            GridLines.Add(line2);
        //        }
        //        resetStart = true;
        //    }
        //}

        //internal void GenerateHorizontalGridLines(ExcelWorksheet ws)
        //{
        //    int addedColumns = PdfWorksheetLayout.AddColumnsForNonWrappedText(ws);
        //    bool resetStart = true;
        //    var startX = 0d;
        //    var startY = 0d;
        //    var endX = 0d;
        //    var endY = 0d;
        //    for (int row = FromRow+1; row <= ToRow; row++)
        //    {
        //        double length = 0;
        //        //string name = "";
        //        //var f = ws.Cells[row, FromCol];
        //        //var start = ChildObjects.Where(x => x.Name == f.Address || x.Name == f.Address + "_m" || x.Name == f.Address + "*").ToList();
        //        for (int col = FromCol; col <= ToCol; col++)
        //        {
        //            var cell = ws.Cells[row, col];
        //            var name = cell.Address;
        //            var layouts = ChildObjects.Where(x => x.Name == cell.Address || x.Name == cell.Address + "_m" || x.Name == cell.Address + "*").ToList();
        //            PdfMergedCellLayout m = null;
        //            PdfCellLayout l = null;
        //            foreach (var layout in layouts)
        //            {
        //                if (layout is PdfMergedCellLayout) m = (PdfMergedCellLayout)layout;
        //                else if (layout is PdfCellLayout) l = (PdfCellLayout)layout;
        //            }
        //            if (l != null)
        //            {
        //                length += l.Size.X;
        //                endX = l.LocalPosition.X + l.Size.X;
        //                endY = l.LocalPosition.Y + l.Size.Y;
        //                if (resetStart)
        //                {
        //                    startX = l.LocalPosition.X;
        //                    startY = l.LocalPosition.Y + l.Size.Y;
        //                    resetStart = false;
        //                }
        //            }
        //            else if (m != null)
        //            {
        //                length += m.Size.X;
        //                endX = m.LocalPosition.X + m.Size.X;
        //                endY = m.LocalPosition.Y + m.Size.Y;
        //                if (resetStart)
        //                {
        //                    startX = m.LocalPosition.X;
        //                    startY = m.LocalPosition.Y + m.Size.Y;
        //                    resetStart = false;
        //                }
        //            }
        //            else if (l == null && m == null)
        //            {
        //                if (length > 0)
        //                {
        //                    length = 0;
        //                    if (startX != 0d)
        //                    {
        //                        var line = new GridLine(startX, startY, endX, endY);
        //                        GridLines.Add(line);
        //                        resetStart = true;
        //                    }
        //                }
        //            }
        //        }
        //        if (startX != 0d)
        //        {
        //            var line2 = new GridLine(startX, startY, endX, endY);
        //            GridLines.Add(line2);
        //        }
        //        resetStart = true;
        //    }
        //}

        //internal void GenerateBorderLines()
        //{
        //    HashSet<double> xCoords = new HashSet<double>();
        //    HashSet<double> yCoords = new HashSet<double>();
        //    // Collect all unique X and Y coordinates.
        //    var cells = ChildObjects.Where(x => x is PdfCellLayout).ToList();
        //    foreach (var c in cells)
        //    {
        //        if (c is PdfMergedCellLayout) continue;
        //        yCoords.Add(c.LocalPosition.Y + c.Size.Y);  // Top
        //        yCoords.Add(c.LocalPosition.Y);             // Bottom
        //        xCoords.Add(c.LocalPosition.X);             // Left
        //        xCoords.Add(c.LocalPosition.X + c.Size.X);  // Right
        //    }
        //    double minX = cells.Where(c => c is not PdfMergedCellLayout).Min(c => c.LocalPosition.X);
        //    double maxX = cells.Where(c => c is not PdfMergedCellLayout).Max(c => c.LocalPosition.X + c.Size.X);
        //    double minY = cells.Where(c => c is not PdfMergedCellLayout).Min(c => c.LocalPosition.Y);
        //    double maxY = cells.Where(c => c is not PdfMergedCellLayout).Max(c => c.LocalPosition.Y + c.Size.Y);
        //    foreach (var x in xCoords.OrderBy(v => v))
        //    {
        //        var line = new GridLine(x, minY, x, maxY);
        //        if (x == minX || x == maxX)
        //            BorderLines.Add(line);
        //    }
        //    foreach (var y in yCoords.OrderBy(v => v))
        //    {
        //        var line = new GridLine(minX, y, maxX, y);
        //        if (y == minY || y == maxY)
        //            BorderLines.Add(line);
        //    }
        //}
    }
}