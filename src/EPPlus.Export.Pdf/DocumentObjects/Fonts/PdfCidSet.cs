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

namespace EPPlus.Export.Pdf.DocumentObjects.Fonts
{
    internal class PdfCidSet : PdfObject
    {
        byte[] CidSet;
        public PdfCidSet(int objectNumber, byte[] cidSet , int version = 0) : base(objectNumber, version)
        {
            CidSet = cidSet;
        }

        internal override string RenderDictionary()
        {
            return $"<< /Length {CidSet.Length} >>\n" + $"stream\n|BINARY DATA|\nendstream";
        }

        internal override void RenderDictionary(BinaryWriter bw)
        {
            WriteAscii(bw, $"<< /Length {CidSet.Length} >>\nstream\n");
            bw.Write(CidSet);
            WriteAscii(bw, "\nendstream");
        }
    }
}
