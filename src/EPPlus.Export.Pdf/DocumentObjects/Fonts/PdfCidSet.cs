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

namespace EPPlus.Export.Pdf.DocumentObjects.Fonts
{
    internal class PdfCidSet : PdfObject
    {
        byte[] CidSet;

        public PdfCidSet(int objectNumber, byte[] cidSet , int version = 0) : base(objectNumber, version)
        {
            CidSet = cidSet;
        }

        protected override void WriteBody(PdfWriter w)
        {
            var body = PdfFlate.Compress(CidSet, w.CompressionLevel);
            if (body.Length < CidSet.Length)
            {
                w.Text($"<< /Filter /FlateDecode /Length {body.Length.ToPdfStringF0()} >>\nstream\n");
            }
            else
            {
                body = CidSet;
                w.Text($"<< /Length {CidSet.Length.ToPdfStringF0()} >>\nstream\n");
            }
            w.Bytes(body);
            w.Text("\nendstream");
        }
    }
}
