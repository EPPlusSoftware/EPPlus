using System.Collections.Generic;

namespace OfficeOpenXml.PDF.PdfLayout
{
    internal class PdfPagesLayout : PdfTransform
    {
        internal ExcelRangeBase Range;

        internal PdfPagesLayout(double x, double y, double height, double width)
            : base(x, y, height, width)
        {
        }

    }
}
