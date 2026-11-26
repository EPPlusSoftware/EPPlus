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
using System.Collections.Generic;
using System.Globalization;

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

        internal void Write(PdfWriter w, long startPosition, int bodyCount)
        {
            StartPosition = startPosition;
            w.Text("xref\n");
            w.Text($"0 {(bodyCount + 1).ToPdfStringF0()}\n");
            w.Text("0000000000 65535 f \n");
            foreach (long pos in positions)
            {
                w.Text(pos.ToString("D10", CultureInfo.InvariantCulture) + " 00000 n \n");
            }
        }
    }
}
