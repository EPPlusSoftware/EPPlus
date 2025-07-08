using OfficeOpenXml.PDF.PdfSettings;
using OfficeOpenXml.PDF.Math;

namespace OfficeOpenXml.PDF.PdfLayout
{
    internal class PdfContentLayout : PdfTransform
    {
        public PdfContentLayout(double x, double y, PdfContentBounds bounds)
        {
            Position = new Vector2(x, y);
            Size = new Vector2(bounds.Width, bounds.Height);
        }
    }
}
