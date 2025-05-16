using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.PDF.PdfObjects
{
    internal class PdfPages : PdfObject
    {
        internal readonly List<int> pageObjectNumbers;

        public PdfPages(int objectNumber, List<int> pageObjectNumbers, int version = 0)
            : base(objectNumber, version)
        {
            this.pageObjectNumbers = pageObjectNumbers.ToList();
        }

        internal override string RenderDictionary()
        {
            var kids = string.Join(" ", pageObjectNumbers.Select(n => $"{n} 0 R").ToArray());
            return $"<< /Type /Pages\n" +
                   $"   /Kids [ {kids} ]\n" +
                   $"   /Count {pageObjectNumbers.Count} >>";
        }
    }
}
