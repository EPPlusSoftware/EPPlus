using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.PDF.PdfObjects
{
    public enum PdfFontSubType
    {
        /*TODO*/Type0,      //Used for Asian fonts
        Type1,      //Used for built-in fonts
        /*TODO*/Type3,      //Custom front
        /*TODO*/TrueType,   //For embedding fonts
    }

    public enum PdfFontEncoding
    {
        None,
        WinAnsiEncoding,
        MacRomanEncoding,
    }

    internal class PdfFont : PdfObject
    {
        private readonly string fontName;
        private readonly PdfFontSubType subType;
        private readonly PdfFontEncoding encoding;

        public PdfFont(int objectNumber, string fontName = "Helvetica", PdfFontSubType subType = PdfFontSubType.Type1, PdfFontEncoding encoding = PdfFontEncoding.WinAnsiEncoding)
            : base(objectNumber, 0)
        {
            this.fontName = fontName;
            this.subType = subType;
            this.encoding = encoding;
        }

        internal override string RenderDictionary()
        {
            var sb = new StringBuilder();
            sb.AppendFormat($"<< /Type /Font\n" +
                   $"   /Subtype /{subType}\n" +
                   $"   /BaseFont /{fontName}");
            if (encoding == PdfFontEncoding.None)
            {
                sb.Append(" >>");
                return sb.ToString();
            }
            sb.AppendFormat($"\n   /Encoding /{encoding} >>");
            return sb.ToString();
        }
    }
}
