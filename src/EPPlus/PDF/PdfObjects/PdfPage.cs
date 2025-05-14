using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.PDF.PdfObjects
{
    internal class PdfPage : PdfObject
    {
        private readonly int parentObjectNumber;
        private readonly int contentObjectNumber;
        private readonly Dictionary<string, int> fontResources;

        public PdfPage(int objectNumber, int parentObjectNumber, int contentObjectNumber, Dictionary<string, int> fontResources, int version = 0)
            : base(objectNumber, version)
        {
            this.parentObjectNumber = parentObjectNumber;
            this.contentObjectNumber = contentObjectNumber;
            this.fontResources = fontResources;
        }

        internal override string RenderDictionary()
        {
            var fontEntries = fontResources.Select(fr => $"/{fr.Key} {fr.Value} 0 R").ToArray();
            var fonts = string.Join(" ", fontEntries);
            return $"<< /Type /Page\n" +
                   $"   /Parent {parentObjectNumber} 0 R\n" +
                   $"   /Resources << /Font << {fonts} >> >>\n" +
                   $"   /MediaBox [0 0 595 842]\n" + // A4 page
                   $"   /Contents {contentObjectNumber} 0 R >>";
        }
    }
}
