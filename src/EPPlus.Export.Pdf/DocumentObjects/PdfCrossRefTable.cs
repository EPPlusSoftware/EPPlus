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
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace EPPlus.Export.Pdf.DocumentObjects
{
    internal class PdfCrossRefTable
    {
        private readonly List<long> positions = new List<long>();
        internal long StartPosition { get; private set; }

        internal void AddPosition(long position)
        {
            positions.Add(position);
        }

        internal string WriteString(int bodyCount)
        {
            StringBuilder sb = new StringBuilder();
            sb.Append("xref\n");
            sb.AppendFormat($"0 {bodyCount + 1}\n");
            sb.Append("0000000000 65535 f \n");
            foreach (long pos in positions)
            {
                sb.AppendFormat(pos.ToString("D10") + " 00000 n \n");
            }
            return sb.ToString();
        }

        internal void Write(BinaryWriter bw, long startPosition, int bodyCount)
        {
            StartPosition = startPosition;
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
