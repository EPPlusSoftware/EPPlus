using EPPlus.Export.Pdf.PdfLayout;
using System.Collections.Generic;

namespace EPPlus.Export.Pdf.PdfCatalog
{
    public enum HeaderFooterType
    {
        First = 0,
        Odd = 1,
        Even = 2
    }

    public enum HeaderFooterAlignment
    {
        Left = 0,
        Center = 1,
        Right = 2
    }

    public enum HeaderFooterSection
    {
        Header = 0,
        Footer = 1
    }

    internal class PdfHeaderFooter
    {
        public HeaderFooterType PageType;
        public HeaderFooterAlignment Alignment;
        public HeaderFooterSection Section;

        public PdfCell Content;
        public bool ContainsPageNumber { get; set; }

        public PdfHeaderFooter(List<PdfTextFormat> textFormats, bool containsPageNumber, HeaderFooterType type, HeaderFooterAlignment alignment, HeaderFooterSection section)
        {
            Content.TextFormats = textFormats;
            ContainsPageNumber = containsPageNumber;
            PageType = type;
            Alignment = alignment;
            Section = section;
        }
    }
}
