using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml.PDF.Math;
using OfficeOpenXml.PDF.PdfSettings;

namespace OfficeOpenXml.PDF.PdfLayout
{
    internal class PdfMergedCellLayout : PdfCellLayout
    {
        public PdfMergedCellLayout() { }

        public PdfMergedCellLayout(ExcelRangeBase cell, PdfPageSettings pageSettings, double x, double y, double width, double height, double scaleX = 1, double scaleY = 1, double rotation = 0, PdfTransform parent = null)
            : base(cell, pageSettings, x, y, width, height, scaleX, scaleY, rotation, parent)
        {
        }
    }
}
