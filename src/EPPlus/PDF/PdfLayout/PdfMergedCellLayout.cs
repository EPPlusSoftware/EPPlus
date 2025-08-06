using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml.PDF.Math;

namespace OfficeOpenXml.PDF.PdfLayout
{
    internal class PdfMergedCellLayout : PdfTransform
    {
        internal double margin = 0.5d;
        public PdfMergedCellLayout() { }

        public PdfMergedCellLayout(object value, double x, double y, double width, double height, double scaleX = 1, double scaleY = 1, double rotation = 0, PdfTransform parent = null)
            : base(x, y, width, height, scaleX, scaleY, rotation, parent)
        {
            //TODO: calculate correct margins.
            //var cellContentLayout = AddChild(new PdfCellContentLayout(value, margin, margin, width - margin * 0.5d, height - margin * 0.5d));
            //cellContentLayout.Name = "Content";
            //cellContentLayout.Z = 2;
        }
    }
}
