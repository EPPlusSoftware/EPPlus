using System.Collections.Generic;

namespace OfficeOpenXml.PDF.PdfPageSettings.PdfPageSizes
{
    internal static class PdfStandardSizes
    {

        public static Dictionary<string, PdfRect> PageFormat = new Dictionary<string, PdfRect>()
        {
            { "A5", new PdfRect{ X = 0d, Y = 0d, Width = 148d, Height = 210d } },
            { "A4", new PdfRect{ X = 0d, Y = 0d, Width = 210d, Height = 297d } },
            { "A3", new PdfRect{ X = 0d, Y = 0d, Width = 297d, Height = 420d } },

            { "B5", new PdfRect{ X = 0d, Y = 0d, Width = 182d, Height = 257d } },
            { "B4", new PdfRect{ X = 0d, Y = 0d, Width = 257d, Height = 364d } },

            { "Letter", new PdfRect{ X = 0d, Y = 0d, Width = 215.9d, Height = 279.4d } },
            { "Legal", new PdfRect{ X = 0d, Y = 0d, Width = 215.9d, Height = 355.6d } },
            { "Statement", new PdfRect{ X = 0d, Y = 0d, Width = 139.7d, Height = 215.9d } },
            { "Executive", new PdfRect{ X = 0d, Y = 0d, Width = 184.2d, Height = 266.7d } },

            { "11x17", new PdfRect{ X = 0d, Y = 0d, Width = 279.4d, Height = 431.8d } },

        };
    }
}
