using System.Collections.Generic;

namespace OfficeOpenXml.PDF.PdfLayout
{
    internal class PdfPageLayout : PdfTransform
    {
        internal List<PdfTransform> ChildObjects = new List<PdfTransform>();
        internal ExcelRangeBase Range;

        internal PdfPageLayout(ExcelRangeBase range)
        {
        }

    }
}
