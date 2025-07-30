using System.Collections.Generic;

namespace OfficeOpenXml.PDF.PdfLayout
{
    internal class PdfPageLayout : PdfTransform
    {
        internal ExcelRangeBase Range;

        public PdfPageLayout(double x, double y, double width, double height)
            :base(x, y, width, height)
        {
        }
    }
}
