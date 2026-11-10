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
using System.Text;

namespace EPPlus.Export.Pdf.DocumentObjects
{
    internal class PdfOutlineItem : PdfObject
    {
        internal string Title;
        internal int Parent, Prev, Next, First, Last, Count;
        internal int DestPageObj;
        internal double DestX, DestY;

        public PdfOutlineItem(int objectNumber, int version = 0)
            : base(objectNumber, version) { }

        private string Body()
        {
            var sb = new StringBuilder();
            sb.Append("<< /Title ").Append(Title.ToPdfTextString());
            sb.Append(" /Parent ").Append(Parent.ToPdfStringF0()).Append(" 0 R");
            if (Prev > 0) sb.Append(" /Prev ").Append(Prev.ToPdfStringF0()).Append(" 0 R");
            if (Next > 0) sb.Append(" /Next ").Append(Next.ToPdfStringF0()).Append(" 0 R");
            if (First > 0) sb.Append(" /First ").Append(First.ToPdfStringF0()).Append(" 0 R");
            if (Last > 0) sb.Append(" /Last ").Append(Last.ToPdfStringF0()).Append(" 0 R");
            if (Count != 0) sb.Append(" /Count ").Append(Count.ToPdfStringF0());
            sb.Append(" /Dest [ ").Append(DestPageObj.ToPdfStringF0()).Append(" 0 R /XYZ ")
              .Append(DestX.ToPdfStringF4()).Append(' ').Append(DestY.ToPdfStringF4()).Append(" null ]");
            sb.Append(" >>");
            return sb.ToString();
        }

        internal override string RenderDictionary() => Body();
        internal override void RenderDictionary(BinaryWriter bw) => WriteAscii(bw, Body());
    }
}