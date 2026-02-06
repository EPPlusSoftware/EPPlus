/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  27/11/2025         EPPlus Software AB           EPPlus 9
 *************************************************************************************************/
using System.Text;

namespace EPPlus.Export.Pdf.PdfObjects.PdfFonts
{
    public enum PdfFontSubType
    {
        Type0,      //Used for embedded fonts
        Type1,      //Used for built-in fonts
        MMType1,
        /*TODO*/Type3,      //Custom front
        /*TODO*/TrueType,   //For embedding fonts
        /*TODO*/CIDFontType0,     //
        /*TODO*/CIDFontType2,     //
    }

    public enum PdfFontEncoding
    {
        None,
        WinAnsiEncoding,
        MacRomanEncoding,
    }

    internal class PdfFont : PdfObject
    {
        private readonly string baseFont;
        private readonly PdfFontSubType subType;
        private readonly PdfFontEncoding encoding;
        private readonly int firstChar;
        private readonly int lastChar;
        private readonly int widthObjectNumber;
        private readonly int fontDescriptorObjectNumber;



        public PdfFont(int objectNumber, string fontName = "Helvetica", PdfFontSubType subType = PdfFontSubType.Type1, int firstChar = -1, int lastChar = -1, int widthObjectNumber = -1, int fontDescObjectNumner = -1, PdfFontEncoding encoding = PdfFontEncoding.WinAnsiEncoding)
            : base(objectNumber, 0)
        {
            baseFont = fontName;
            this.subType = subType;
            this.encoding = encoding;
            this.firstChar = firstChar;
            this.lastChar = lastChar;
            this.widthObjectNumber = widthObjectNumber;
            fontDescriptorObjectNumber = fontDescObjectNumner;
        }

        internal override string RenderDictionary()
        {
            var sb = new StringBuilder();
            sb.AppendFormat($"<< /Type /Font\n" +
                            $"   /Subtype /{subType}\n" +
                            $"   /BaseFont /{baseFont.Replace(" ", "")}");
            if (encoding == PdfFontEncoding.None)
            {
                sb.Append(" >>");
                return sb.ToString();
            }
            else
            {
                sb.Append("\n");
            }
            if (firstChar > -1)
            {
                sb.AppendFormat($"   /FirstChar {firstChar}\n");
            }
            if(lastChar > -1)
            {
                sb.AppendFormat($"   /LastChar {lastChar}\n");
            }
            if (widthObjectNumber > -1)
            {
                sb.AppendFormat($"   /Widths {widthObjectNumber } 0 R\n");
            }
            if (fontDescriptorObjectNumber > -1)
            {
                sb.AppendFormat($"   /FontDescriptor {fontDescriptorObjectNumber} 0 R\n");
            }
            sb.AppendFormat($"   /Encoding /{encoding} >>");
            return sb.ToString();
        }
    }
}
