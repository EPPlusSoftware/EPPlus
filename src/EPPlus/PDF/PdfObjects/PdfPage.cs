using OfficeOpenXml.PDF.Pdfhelpers;
using OfficeOpenXml.PDF.PdfResources;
using OfficeOpenXml.PDF.PdfSettings.PdfPageSizes;
using System.Collections.Generic;
using System.Linq;

namespace OfficeOpenXml.PDF.PdfObjects
{
    internal class PdfPage : PdfObject
    {
        private readonly int parentObjectNumber;
        internal readonly List<int> contentObjectNumbers;
        PdfDictionaries dictionaries;
        internal PdfPageSize Size;

        public PdfPage(int objectNumber, int parentObjectNumber, List<int> contentObjectNumbers, PdfPageSize size, PdfDictionaries dictionaries, int version = 0)
            : base(objectNumber, version)
        {
            this.parentObjectNumber = parentObjectNumber;
            this.contentObjectNumbers = contentObjectNumbers;
            this.dictionaries = dictionaries;
            Size = size;
        }

        internal override string RenderDictionary()
        {
            var fontEntries = dictionaries.Fonts.Select(f => $"/{f.Value.Label} {f.Value.fontObjectNumber} 0 R").ToArray();
            var fonts = string.Join(" ", fontEntries);
            var patternEntries = dictionaries.Patterns.Select(p => $"/{p.Value.Label} {p.Value.shadingPatternobjectNumber} 0 R").ToArray();
            var patterns = string.Join(" ", patternEntries);
            var shadingEntries = dictionaries.Shadings.Select(s => $"/{s.Value.Label} {s.Value.shadingObjectNumber} 0 R").ToArray();
            var shadings = string.Join(" ", shadingEntries);
            var contentEntries = contentObjectNumbers.Select(con => $"{con} 0 R").ToArray();
            return $"<< /Type /Page\n" +
                   $"   /Parent {parentObjectNumber} 0 R\n" +
                   $"   /Resources << /Font << {fonts} >>\n" +
                   $"                 /Pattern << {patterns} >>\n" +
                   $"                 /Shading << {shadings} >> >>\n" +
                   $"   /MediaBox [ 0 0 {Size.WidthPu.ToPdfString()} {Size.HeightPu.ToPdfString()} ]\n" +
                   $"   /Contents [ {string.Join(" ", contentEntries)} ] >>";
        }
    }
}
