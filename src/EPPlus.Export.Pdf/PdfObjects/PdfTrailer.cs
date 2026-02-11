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
using System.IO;
using System.Text;

namespace EPPlus.Export.Pdf.PdfObjects
{
    internal static class PdfTrailer
    {
        internal static string WriteString(int bodyCount, int catalogObjectNumber, int infoObjectNumber, long crossRefStartPosition)
        {
            StringBuilder sb = new StringBuilder();
            sb.AppendFormat("trailer\n");
            sb.AppendFormat($"<< /Size {bodyCount + 1}\n" +
                            $"   /Root {catalogObjectNumber} 0 R\n" +
                            $"   /Info {infoObjectNumber} 0 R >>\n");
            sb.AppendFormat("startxref\n");
            sb.AppendFormat($"{crossRefStartPosition}\n");
            sb.AppendFormat("%%EOF\n");
            return sb.ToString();
        }

        internal static void Write(BinaryWriter bw, int bodyCount, int catalogObjectNumber, int infoObjectNumber, long crossRefStartPosition)
        {
            bw.Write(Encoding.ASCII.GetBytes($"trailer\n" +
                                             $"<< /Size {bodyCount + 1}\n" +
                                             $"   /Root {catalogObjectNumber} 0 R\n" +
                                             $"   /Info {infoObjectNumber} 0 R >>\n" +
                                             $"startxref\n" +
                                             $"{crossRefStartPosition}\n" +
                                             $"%%EOF\n"));
        }
    }
}
