using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.PDF
{
    internal class PdfCrossRefTable
    {
        private readonly List<long> positions = new List<long>();
        internal long StartPosition { get; private set; }


        internal void AddPosition(long position)
        {
            positions.Add(position);
        }

        internal void Write(BinaryWriter bw, long startPosition, int bodyCount)
        {
            this.StartPosition = startPosition;
            bw.Write(Encoding.ASCII.GetBytes("xref\n"));
            bw.Write(Encoding.ASCII.GetBytes($"0 {bodyCount + 1}\n"));
            bw.Write(Encoding.ASCII.GetBytes("0000000000 65535 f \n"));
            foreach (long pos in positions)
            {
                bw.Write(Encoding.ASCII.GetBytes(pos.ToString("D10") + " 00000 n \n"));
            }
        }
    }
}
