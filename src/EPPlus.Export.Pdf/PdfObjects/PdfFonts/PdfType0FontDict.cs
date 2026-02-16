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
using System.IO;
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
            BaseFont = string.Concat(basefont.Where(c => !char.IsWhiteSpace(c)));
            Encoding = encoding;
            DescendantFontsObjectNumbers = descendantFontsObjectNumbers;
            ToUnicodeObjectNumber = toUnicodeObjectNumber;
        }

        internal override string RenderDictionary()
        {
            //var DescendantFonts = string.Join(" ", DescendantFontsObjectNumbers.Select(w => ($"{w} 0 R ").ToString()).ToArray());
            var sb = new StringBuilder();
            sb.AppendFormat($"<<  /Type /Font\n" +
                            $"    /Subtype /Type0\n" +
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

        internal override void RenderDictionary(BinaryWriter bw)
        {
            //var DescendantFonts = string.Join(" ", DescendantFontsObjectNumbers.Select(w => ($"{w} 0 R ").ToString()).ToArray());
            var sb = new StringBuilder();
            sb.AppendFormat($"<<  /Type /Font\n" +
                            $"    /Subtype /Type0\n" +
                            $"    /BaseFont /{BaseFont}\n" +
                            $"    /Encoding /{Encoding}\n" +
                            $"    /DescendantFonts [{DescendantFontsObjectNumbers} 0 R]");
            if (ToUnicodeObjectNumber > 0)
            {
                sb.AppendFormat($"\n    /ToUnicode {ToUnicodeObjectNumber} 0 R");
            }
            sb.Append(" >>");
            WriteAscii(bw, sb.ToString());
        }
    }
}
