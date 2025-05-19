using System.Collections.Generic;

namespace OfficeOpenXml.PDF.PdfPageSettings.PdfPageSizes
{
    internal static class PdfStandardSizes
    {

        public static Dictionary<string, PdfRect> PageFormat = new Dictionary<string, PdfRect>()
        {
            { "A5", new PdfRect{ X = 0f, Y = 0f, Width = 148f, Height = 210f } },
            { "A4", new PdfRect{ X = 0f, Y = 0f, Width = 210f, Height = 297f } },
            { "A3", new PdfRect{ X = 0f, Y = 0f, Width = 297f, Height = 420f } },

            { "B5", new PdfRect{ X = 0f, Y = 0f, Width = 182f, Height = 257f } },
            { "B4", new PdfRect{ X = 0f, Y = 0f, Width = 257f, Height = 364f } },

            { "Letter", new PdfRect{ X = 0f, Y = 0f, Width = 215.9f, Height = 279.4f } },
            { "Legal", new PdfRect{ X = 0f, Y = 0f, Width = 215.9f, Height = 355.6f } },
            { "Statement", new PdfRect{ X = 0f, Y = 0f, Width = 139.7f, Height = 215.9f } },
            { "Executive", new PdfRect{ X = 0f, Y = 0f, Width = 184.2f, Height = 266.7f } },

            { "11x17", new PdfRect{ X = 0f, Y = 0f, Width = 279.4f, Height = 431.8f } },

        };
    }
}
