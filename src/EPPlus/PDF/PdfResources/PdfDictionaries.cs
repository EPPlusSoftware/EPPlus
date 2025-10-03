using System.Collections.Generic;

namespace OfficeOpenXml.PDF.PdfResources
{
    internal class PdfDictionaries
    {
        internal readonly Dictionary<string, PdfFontResource> Fonts = new Dictionary<string, PdfFontResource>();
        internal readonly Dictionary<string, PdfPatternResource> Patterns = new Dictionary<string, PdfPatternResource>();
        internal readonly Dictionary<string, PdfShadingResource> Shadings = new Dictionary<string, PdfShadingResource>();
    }
}
