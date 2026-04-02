using EPPlus.Export.Pdf.PdfResources;
using EPPlus.Export.Pdf.PdfSettings;
using OfficeOpenXml;
using System.Collections.Generic;

namespace EPPlus.Export.Pdf.PdfCatalog
{
    internal class PdfHeaderFooterCollection
    {
        public List<PdfHeaderFooter> PdfHeaderFooterEntries = new List<PdfHeaderFooter>();

        public PdfHeaderFooterCollection(PdfPageSettings pageSettings, PdfDictionaries dictionaries, PdfWorksheet pdfSheet, ExcelHeaderFooter headerFooter)
        {
            //First Header
            PdfHeaderFooterEntries.Add(PdfTextMap.GetTextFormats(pageSettings, dictionaries, pdfSheet.Worksheet, headerFooter.FirstHeader.LeftAligned, HeaderFooterType.First, HeaderFooterAlignment.Left, HeaderFooterSection.Header));
            PdfHeaderFooterEntries.Add(PdfTextMap.GetTextFormats(pageSettings, dictionaries, pdfSheet.Worksheet, headerFooter.FirstHeader.Centered, HeaderFooterType.First, HeaderFooterAlignment.Center, HeaderFooterSection.Header));
            PdfHeaderFooterEntries.Add(PdfTextMap.GetTextFormats(pageSettings, dictionaries, pdfSheet.Worksheet, headerFooter.FirstHeader.RightAligned, HeaderFooterType.First, HeaderFooterAlignment.Right, HeaderFooterSection.Header));
            //First Footer
            PdfHeaderFooterEntries.Add(PdfTextMap.GetTextFormats(pageSettings, dictionaries, pdfSheet.Worksheet, headerFooter.FirstFooter.LeftAligned, HeaderFooterType.First, HeaderFooterAlignment.Left, HeaderFooterSection.Footer));
            PdfHeaderFooterEntries.Add(PdfTextMap.GetTextFormats(pageSettings, dictionaries, pdfSheet.Worksheet, headerFooter.FirstFooter.Centered, HeaderFooterType.First, HeaderFooterAlignment.Center, HeaderFooterSection.Footer));
            PdfHeaderFooterEntries.Add(PdfTextMap.GetTextFormats(pageSettings, dictionaries, pdfSheet.Worksheet, headerFooter.FirstFooter.RightAligned, HeaderFooterType.First, HeaderFooterAlignment.Right, HeaderFooterSection.Footer));
            //Odd Header
            PdfHeaderFooterEntries.Add(PdfTextMap.GetTextFormats(pageSettings, dictionaries, pdfSheet.Worksheet, headerFooter.OddHeader.LeftAligned, HeaderFooterType.Odd, HeaderFooterAlignment.Left, HeaderFooterSection.Header));
            PdfHeaderFooterEntries.Add(PdfTextMap.GetTextFormats(pageSettings, dictionaries, pdfSheet.Worksheet, headerFooter.OddHeader.Centered, HeaderFooterType.Odd, HeaderFooterAlignment.Center, HeaderFooterSection.Header));
            PdfHeaderFooterEntries.Add(PdfTextMap.GetTextFormats(pageSettings, dictionaries, pdfSheet.Worksheet, headerFooter.OddHeader.RightAligned, HeaderFooterType.Odd, HeaderFooterAlignment.Right, HeaderFooterSection.Header));
            //Odd Footer
            PdfHeaderFooterEntries.Add(PdfTextMap.GetTextFormats(pageSettings, dictionaries, pdfSheet.Worksheet, headerFooter.OddFooter.LeftAligned, HeaderFooterType.Odd, HeaderFooterAlignment.Left, HeaderFooterSection.Footer));
            PdfHeaderFooterEntries.Add(PdfTextMap.GetTextFormats(pageSettings, dictionaries, pdfSheet.Worksheet, headerFooter.OddFooter.Centered, HeaderFooterType.Odd, HeaderFooterAlignment.Center, HeaderFooterSection.Footer));
            PdfHeaderFooterEntries.Add(PdfTextMap.GetTextFormats(pageSettings, dictionaries, pdfSheet.Worksheet, headerFooter.OddFooter.RightAligned, HeaderFooterType.Odd, HeaderFooterAlignment.Right, HeaderFooterSection.Footer));
            //Even Header
            PdfHeaderFooterEntries.Add(PdfTextMap.GetTextFormats(pageSettings, dictionaries, pdfSheet.Worksheet, headerFooter.EvenHeader.LeftAligned, HeaderFooterType.Even, HeaderFooterAlignment.Left, HeaderFooterSection.Header));
            PdfHeaderFooterEntries.Add(PdfTextMap.GetTextFormats(pageSettings, dictionaries, pdfSheet.Worksheet, headerFooter.EvenHeader.Centered, HeaderFooterType.Even, HeaderFooterAlignment.Center, HeaderFooterSection.Header));
            PdfHeaderFooterEntries.Add(PdfTextMap.GetTextFormats(pageSettings, dictionaries, pdfSheet.Worksheet, headerFooter.EvenHeader.RightAligned, HeaderFooterType.Even, HeaderFooterAlignment.Right, HeaderFooterSection.Header));
            //Even Footer
            PdfHeaderFooterEntries.Add(PdfTextMap.GetTextFormats(pageSettings, dictionaries, pdfSheet.Worksheet, headerFooter.EvenFooter.LeftAligned, HeaderFooterType.Even, HeaderFooterAlignment.Left, HeaderFooterSection.Footer));
            PdfHeaderFooterEntries.Add(PdfTextMap.GetTextFormats(pageSettings, dictionaries, pdfSheet.Worksheet, headerFooter.EvenFooter.Centered, HeaderFooterType.Even, HeaderFooterAlignment.Center, HeaderFooterSection.Footer));
            PdfHeaderFooterEntries.Add(PdfTextMap.GetTextFormats(pageSettings, dictionaries, pdfSheet.Worksheet, headerFooter.EvenFooter.RightAligned, HeaderFooterType.Even, HeaderFooterAlignment.Right, HeaderFooterSection.Footer));
        }
    }
}
