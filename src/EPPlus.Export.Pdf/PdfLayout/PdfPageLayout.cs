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
using System.Collections.Generic;
using System.Linq;

namespace EPPlus.Export.Pdf.PdfLayout
{
    internal class PdfPageLayout : PdfTransform
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
                yCoords.Add(c.LocalPosition.Y + c.Size.Y);             // Top
                yCoords.Add(c.LocalPosition.Y);  // Bottom
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
    }
}
