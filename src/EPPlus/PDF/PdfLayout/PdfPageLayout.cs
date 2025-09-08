using OfficeOpenXml.PDF.PdfSettings;
using System.Collections.Generic;
using System.Linq;

namespace OfficeOpenXml.PDF.PdfLayout
{
    internal class PdfPageLayout : PdfTransform
    {
        internal List<GridLine> GridLines = new List<GridLine>();
        internal List<GridLine> BorderLines = new List<GridLine>();

        public PdfPageLayout(double x, double y, double width, double height)
            :base(x, y, width, height)
        {
        }

        internal void GenerateGridLines(PdfPageSettings pageSettings, ExcelWorksheet ws)
        {
            HashSet<double> xCoords = new HashSet<double>();
            HashSet<double> yCoords = new HashSet<double>();

            // Collect all unique X and Y coordinates
            var cells = ChildObjects.Where(x => x is PdfCellLayout).ToList();
            foreach (var c in cells)
            {
                if (c is PdfMergedCellLayout) continue;
                xCoords.Add(c.LocalPosition.X);              // left
                xCoords.Add(c.LocalPosition.X + c.Size.X);    // right
                yCoords.Add(c.LocalPosition.Y);              // top
                yCoords.Add(c.LocalPosition.Y - c.Size.Y);   // bottom
            }

            double minX = cells.Where(c => c is not PdfMergedCellLayout).Min(c => c.LocalPosition.X);
            double maxX = cells.Where(c => c is not PdfMergedCellLayout).Max(c => c.LocalPosition.X + c.Size.X);
            double minY = cells.Where(c => c is not PdfMergedCellLayout).Min(c => c.LocalPosition.Y - c.Size.Y);
            double maxY = cells.Where(c => c is not PdfMergedCellLayout).Max(c => c.LocalPosition.Y);

            List<GridLine> lines = new List<GridLine>();

            foreach (var x in xCoords.OrderBy(v => v))
            {
                var line = new GridLine(x, minY, x, maxY);
                if (x == minX || x == maxX)
                    BorderLines.Add(line);
                else
                    GridLines.Add(line);
            }

            // Horizontal lines
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
