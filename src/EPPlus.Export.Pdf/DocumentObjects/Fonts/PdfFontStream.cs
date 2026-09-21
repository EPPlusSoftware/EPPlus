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
using EPPlus.Export.Pdf.Helpers;
using EPPlus.Fonts.OpenType;
using System.IO;
using System.Text;

namespace EPPlus.Export.Pdf.DocumentObjects.Fonts
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
            return $"<< /Length {fontBytes.Length.ToPdfStringF0()} /Length1 {fontBytes.Length.ToPdfStringF0()} >>\n" + $"stream\n|BINARY DATA|\nendstream";
        }

        internal override void RenderDictionary(BinaryWriter bw)
        {
            var fontBytes = FontData.Serialize();
            var body = PdfFlate.Compress(fontBytes);
            WriteAscii(bw, $"<< /Filter /FlateDecode /Length {body.Length.ToPdfStringF0()} /Length1 {fontBytes.Length.ToPdfStringF0()} >>\nstream\n");
            bw.Write(body);
            WriteAscii(bw, "\nendstream");
        }
    }
}
