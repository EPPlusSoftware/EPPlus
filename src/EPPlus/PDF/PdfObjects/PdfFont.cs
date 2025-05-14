using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.PDF.PdfObjects
{
    internal enum PdfFontSubType
    {
        /*TODO*/Type0,      //Used for Asian fonts
        Type1,      //Used for built-in fonts
        /*TODO*/Type3,      //Custom front
        /*TODO*/TrueType,   //For embedding fonts
    }

    internal class PdfFont : PdfObject
    {
        private readonly string fontName;
        private readonly PdfFontSubType subType;
        private readonly string encoding;

        public PdfFont(int objectNumber, string fontName = "Helvetica", PdfFontSubType subType = PdfFontSubType.Type1, string encoding = "WinAnsiEncoding")
            : base(objectNumber, 0)
        {
            this.fontName = fontName;
            this.subType = subType;
            this.encoding = encoding;
        }

        internal override string RenderDictionary()
        {
            return $"<< /Type /Font\n" +
                   $"   /Subtype /{subType}\n" +
                   $"   /BaseFont /{fontName}\n" +
                   $"   /Encoding /{encoding} >>";
        }
    }
}
