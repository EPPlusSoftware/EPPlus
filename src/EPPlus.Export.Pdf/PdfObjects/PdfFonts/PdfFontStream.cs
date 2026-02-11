using EPPlus.Fonts.OpenType;
using System.IO;
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
            var fontData = Encoding.ASCII.GetString(fontBytes);
            return $"<< /Length {fontBytes.Length} >>\n" + $"stream\n|BINARY DATA|\nendstream";
        }

        internal override void RenderDictionary(BinaryWriter bw)
        {
            var fontBytes = FontData.Serialize();
            WriteAscii(bw, $"<< /Length {fontBytes.Length} >>\nstream\n");
            bw.Write(fontBytes);
            WriteAscii(bw, "\nendstream");
        }
    }
}
