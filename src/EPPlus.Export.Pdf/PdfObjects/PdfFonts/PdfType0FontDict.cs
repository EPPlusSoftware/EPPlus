using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;

namespace EPPlus.Export.Pdf.PdfObjects.PdfFonts
{
    internal class PdfType0FontDict : PdfObject
    {
        private readonly string BaseFont;
        private readonly string Encoding;
        private readonly int DescendantFontsObjectNumbers;

        private readonly int ToUnicodeObjectNumber;

        public PdfType0FontDict(int objectNumber, string basefont, string encoding, int descendantFontsObjectNumbers, int toUnicodeObjectNumber = -1, int version = 0)
            : base(objectNumber, version)
        {
            BaseFont = basefont;
            Encoding = encoding;
            DescendantFontsObjectNumbers = descendantFontsObjectNumbers;
            ToUnicodeObjectNumber = toUnicodeObjectNumber;
        }

        internal override string RenderDictionary()
        {
            //var DescendantFonts = string.Join(" ", DescendantFontsObjectNumbers.Select(w => ($"{w} 0 R ").ToString()).ToArray());
            var sb = new StringBuilder();
            sb.AppendFormat($"<<  /Type /Font\n" +
                            $"    /SubType /Type0\n" +
                            $"    /BaseFont /{BaseFont}\n" +
                            $"    /Encoding /{Encoding}\n" +
                            $"    /DescendantFonts [{DescendantFontsObjectNumbers} 0 R]");
            if (ToUnicodeObjectNumber > 0)
            {
                sb.AppendFormat($"\n    /ToUnicode {ToUnicodeObjectNumber} 0 R");
            }
            sb.Append(" >>");
            return sb.ToString();
        }
    }
}
