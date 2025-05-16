using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.PDF.PdfObjects
{
    internal class PdfPage : PdfObject
    {
        private readonly int parentObjectNumber;
        internal readonly List<int> contentObjectNumbers;
        internal readonly Dictionary<int, string> fontResources;

        public PdfPage(int objectNumber, int parentObjectNumber, List<int> contentObjectNumbers, Dictionary<int, string> fontResources, int version = 0)
            : base(objectNumber, version)
        {
            this.parentObjectNumber = parentObjectNumber;
            this.contentObjectNumbers = contentObjectNumbers;
            this.fontResources = fontResources;
        }

        internal override string RenderDictionary()
        {
            var fontEntries = fontResources.Select(fr => $"/{fr.Value} {fr.Key} 0 R").ToArray();
            var contentEntries = contentObjectNumbers.Select(con => $"{con} 0 R").ToArray();
            var fonts = string.Join(" ", fontEntries);
            return $"<< /Type /Page\n" +
                   $"   /Parent {parentObjectNumber} 0 R\n" +
                   $"   /Resources << /Font << {fonts} >> >>\n" +
                   $"   /MediaBox [0 0 595 842]\n" + // A4 page
                   $"   /Contents [ {string.Join(" ", contentEntries)} ]>>";
        }
    }
}
