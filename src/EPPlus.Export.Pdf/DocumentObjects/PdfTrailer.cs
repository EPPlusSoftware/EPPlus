/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  10/07/2025         EPPlus Software AB           EPPlus.Fonts.OpenType 1.0
 *************************************************************************************************/
using EPPlus.Export.Pdf.Helpers;

namespace EPPlus.Export.Pdf.DocumentObjects
{
    internal static class PdfTrailer
    {
        internal static void Write(PdfWriter w, int bodyCount, int catalogObjectNumber, int infoObjectNumber, long crossRefStartPosition)
        {
            w.Text("trailer\n");
            w.Text($"<< /Size {(bodyCount + 1).ToPdfStringF0()}\n" +
                   $"   /Root {catalogObjectNumber.ToPdfStringF0()} 0 R\n" +
                   $"   /Info {infoObjectNumber.ToPdfStringF0()} 0 R >>\n");
            w.Text("startxref\n");
            w.Text($"{crossRefStartPosition.ToPdfStringF0()}\n");
            w.Text("%%EOF\n");
        }
    }
}
