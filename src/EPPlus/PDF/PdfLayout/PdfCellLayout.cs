using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml.PDF.Math;

namespace OfficeOpenXml.PDF.PdfLayout
{
    internal class PdfCellLayout : PdfTransform
    {
        internal PdfTransform Value;

        public PdfCellLayout() { }

        public PdfCellLayout(PdfTransform value, double x, double y, double width, double height)
            :base(x, y, width, height)
        {
            Value = value;
        }
    }
}
