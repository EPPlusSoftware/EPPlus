using OfficeOpenXml.PDF.PdfFontData;
using OfficeOpenXml.PDF.Pdfhelpers;
using OfficeOpenXml.PDF.PdfSettings.PdfPageSizes;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.PDF.PdfObjects
{
    internal class PdfPage : PdfObject
    {
        private readonly int parentObjectNumber;
        internal readonly List<int> contentObjectNumbers;
        internal readonly Dictionary<string, PdfFontResource> fontResources;
        internal PdfPageSize Size;

        public PdfPage(int objectNumber, int parentObjectNumber, List<int> contentObjectNumbers, PdfPageSize size, Dictionary<string, PdfFontResource> fontResources, int version = 0)
            : base(objectNumber, version)
        {
            this.parentObjectNumber = parentObjectNumber;
            this.contentObjectNumbers = contentObjectNumbers;
            this.fontResources = fontResources;
            Size = size;
        }

        internal override string RenderDictionary()
        {
            var fontEntries = fontResources.Select(fr => $"/{fr.Value.Label} {fr.Value.fontObjectNumber} 0 R").ToArray();
            var contentEntries = contentObjectNumbers.Select(con => $"{con} 0 R").ToArray();
            var fonts = string.Join(" ", fontEntries);
            return $"<< /Type /Page\n" +
                   $"   /Parent {parentObjectNumber} 0 R\n" +
                   $"   /Resources << /Font << {fonts} >> >>\n" +
                   $"   /MediaBox [0 0 {Size.WidthPu.ToPdfString()} {Size.HeightPu.ToPdfString()}]\n" +
                   $"   /Contents [ {string.Join(" ", contentEntries)} ] >>";
        }
    }
}
