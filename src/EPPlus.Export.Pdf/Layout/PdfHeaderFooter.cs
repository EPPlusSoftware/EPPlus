/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  27/11/2025         EPPlus Software AB           EPPlus 9
 *************************************************************************************************/
using EPPlus.Fonts.OpenType.Integration;
using System.Collections.Generic;

namespace EPPlus.Export.Pdf.Layout
{
    internal enum HeaderFooterType
    {
        First = 0,
        Odd = 1,
        Even = 2
    }

    internal enum HeaderFooterAlignment
    {
        Left = 0,
        Center = 1,
        Right = 2
    }

    internal enum HeaderFooterSection
    {
        Header = 0,
        Footer = 1
    }

    internal class PdfHeaderFooter
    {
        public HeaderFooterType PageType;
        public HeaderFooterAlignment Alignment;
        public HeaderFooterSection Section;
        public PdfCellBase Content;
        public List<int> NumberOfPagesIndexes = new List<int>();
        public List<int> PageNumberIndexes = new List<int>();

        public byte[] ImageBytes;
        public double ImageWidth;
        public double ImageHeight;
        public int ImageFragmentIndex = -1;
        public bool HasImage => ImageBytes != null;

        public PdfHeaderFooter(List<TextFragment> textFormats, List<int> pageNumberIndexes, List<int> numberOfPagesIndexes, HeaderFooterType type, HeaderFooterAlignment alignment, HeaderFooterSection section)
        {
            Content = new PdfCellBase();
            Content.TextFragments = textFormats;
            PageNumberIndexes = pageNumberIndexes;
            NumberOfPagesIndexes = numberOfPagesIndexes;
            PageType = type;
            Alignment = alignment;
            Section = section;
        }
    }

    internal class PdfHeaderFooterImage
    {
        public HeaderFooterType PageType;
        public HeaderFooterAlignment Alignment;
        public HeaderFooterSection Section;
        public byte[] ImageBytes;
        public double Width;
        public double Height;

        public PdfHeaderFooterImage(byte[] imageBytes, double width, double height, HeaderFooterType type, HeaderFooterAlignment alignment, HeaderFooterSection section)
        {
            ImageBytes = imageBytes;
            Width = width;
            Height = height;
            PageType = type;
            Alignment = alignment;
            Section = section; }
    }
}
