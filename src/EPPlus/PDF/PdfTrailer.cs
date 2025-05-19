using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.PDF
{
    internal static class PdfTrailer
    {
        internal static void Write(BinaryWriter bw, int bodyCount, int catalogObjectNumber, long crossRefStartPosition)
        {
            bw.Write(Encoding.ASCII.GetBytes("trailer\n"));
            bw.Write(Encoding.ASCII.GetBytes($"<< /Size {bodyCount + 1} /Root {catalogObjectNumber} 0 R >>\n"));
            bw.Write(Encoding.ASCII.GetBytes("startxref\n"));
            bw.Write(Encoding.ASCII.GetBytes($"{crossRefStartPosition}\n"));
            bw.Write(Encoding.ASCII.GetBytes("%%EOF\n"));
        }
    }
}
