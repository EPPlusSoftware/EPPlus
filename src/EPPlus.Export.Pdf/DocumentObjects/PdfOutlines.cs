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
using System.IO;

namespace EPPlus.Export.Pdf.DocumentObjects
{
    internal class PdfOutlines : PdfObject
    {
        internal int First, Last, Count;

        public PdfOutlines(int objectNumber, int version = 0)
            : base(objectNumber, version) { }

        private string Body() =>
            $"<< /Type /Outlines" +
            $" /First {First.ToPdfStringF0()} 0 R" +
            $" /Last {Last.ToPdfStringF0()} 0 R" +
            $" /Count {Count.ToPdfStringF0()} >>";

        internal override string RenderDictionary() => Body();
        internal override void RenderDictionary(BinaryWriter bw) => WriteAscii(bw, Body());
    }
}