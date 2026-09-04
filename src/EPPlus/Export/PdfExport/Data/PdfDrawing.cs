using OfficeOpenXml.Drawing;

namespace OfficeOpenXml.Export.PdfExport.Data
{
    internal class PdfDrawing
    {
        public ExcelPicture Picture { get; }
        public byte[] ImageBytes => Picture.Image.ImageBytes;   // raw JPEG stream, embeds verbatim later
        public ePictureType PictureType => Picture.Image.Type.Value;
        public PdfDrawing(ExcelPicture picture) { Picture = picture; }
    }
}
