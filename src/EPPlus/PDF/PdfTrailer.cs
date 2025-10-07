using System.IO;
using System.Text;

namespace OfficeOpenXml.PDF
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
            bw.Write(Encoding.ASCII.GetBytes("trailer\n"));
            bw.Write(Encoding.ASCII.GetBytes($"<< /Size {bodyCount + 1}\n" +
                                             $"   /Root {catalogObjectNumber} 0 R\n" +
                                             $"   /Info {infoObjectNumber} 0 R >>\n"));
            bw.Write(Encoding.ASCII.GetBytes("startxref\n"));
            bw.Write(Encoding.ASCII.GetBytes($"{crossRefStartPosition}\n"));
            bw.Write(Encoding.ASCII.GetBytes("%%EOF\n"));
        }
    }
}
