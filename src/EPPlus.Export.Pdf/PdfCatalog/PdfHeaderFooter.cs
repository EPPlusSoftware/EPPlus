using EPPlus.Export.Pdf.PdfLayout;
using EPPlus.Fonts.OpenType.Integration;
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

        public List<int> NumberOfPagesIndexes = new List<int>();
        public List<int> PageNumberIndexes = new List<int>();

        public PdfHeaderFooter(List<TextFragment> textFormats, List<int> pageNumberIndexes, List<int> numberOfPagesIndexes, HeaderFooterType type, HeaderFooterAlignment alignment, HeaderFooterSection section)
        {
            Content = new PdfCell();
            Content.TextFragments = textFormats;
            PageNumberIndexes = pageNumberIndexes;
            NumberOfPagesIndexes = numberOfPagesIndexes;
            PageType = type;
            Alignment = alignment;
            Section = section;
        }
    }
}
