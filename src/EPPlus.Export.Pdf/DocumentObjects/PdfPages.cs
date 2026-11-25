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
using System.Linq;

namespace EPPlus.Export.Pdf.DocumentObjects
{
    internal class PdfPages : PdfObject
    {
        internal readonly List<int> pageObjectNumbers;

        public PdfPages(int objectNumber, List<int> pageObjectNumbers, int version = 0)
            : base(objectNumber, version)
        {
            this.pageObjectNumbers = pageObjectNumbers.ToList();
        }

        private string Body()
        {
            var kids = string.Join(" ", pageObjectNumbers.Select(n => $"{n.ToPdfStringF0()} 0 R").ToArray());
            return $"<< /Type /Pages\n" +
                   $"   /Kids [ {kids} ]\n" +
                   $"   /Count {pageObjectNumbers.Count.ToPdfStringF0()} >>";
        }

        protected override void WriteBody(PdfWriter w) => w.Text(Body());
    }
}
