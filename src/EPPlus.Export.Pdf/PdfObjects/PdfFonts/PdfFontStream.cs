using EPPlus.Fonts.OpenType;
using System.Text;

namespace EPPlus.Export.Pdf.PdfObjects.PdfFonts
{
    internal class PdfFontStream : PdfObject
    {
        private readonly OpenTypeFont FontData;

        public PdfFontStream(int objectNumber, OpenTypeFont fontData, int version = 0) : base(objectNumber, version)
        {
            FontData = fontData;
        }

        internal override string RenderDictionary()
        {
            var fontBytes = FontData.Serialize();
            return $"<< /Length {fontBytes.Length} >>\n" + $"stream\n{fontBytes}endstream";
        }
    }
}
