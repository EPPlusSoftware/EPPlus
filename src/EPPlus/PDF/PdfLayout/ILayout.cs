using OfficeOpenXml.PDF.PdfSettings;

namespace OfficeOpenXml.PDF.PdfLayout
{
    internal interface ILayout
    {
        public void ConvertCoordinates(PdfPageSettings pageSettings);
    }
}
