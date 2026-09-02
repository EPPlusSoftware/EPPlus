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
using EPPlus.Export.Pdf.Layout;
using EPPlus.Export.Pdf.Resources;
using EPPlus.Export.Pdf.Settings;
using OfficeOpenXml.Drawing.Vml;
using OfficeOpenXml.Export.PdfExport.Data;
using System.Collections.Generic;
using System.Linq;

namespace OfficeOpenXml.Export.PdfExport.TextMapping
{
    internal class PdfHeaderFooterCollection
    {
        public List<PdfHeaderFooter> PdfHeaderFooterEntries = new List<PdfHeaderFooter>();
        public List<PdfHeaderFooterImage> PdfHeaderFooterImages = new List<PdfHeaderFooterImage>();
        public bool ScaleWithDocument = false;
        public bool AlignWithMargins = false;
        public bool HasFirstPage = false;
        public bool HasOddEvenPages = false;

        public PdfHeaderFooterCollection(PdfPageSettings pageSettings, PdfDictionaries dictionaries, PdfWorksheet pdfSheet, ExcelHeaderFooter headerFooter)
        {
            bool differentFirst = pdfSheet.Worksheet.HeaderFooter.differentFirst;
            bool differentOddEven = pdfSheet.Worksheet.HeaderFooter.differentOddEven;
            HasFirstPage = differentFirst;
            HasOddEvenPages = differentOddEven;
            bool AlignWithMargins = pdfSheet.Worksheet.HeaderFooter.AlignWithMargins;
            bool ScaleWithDocument = pdfSheet.Worksheet.HeaderFooter.ScaleWithDocument;
            PdfHeaderFooter entry = null;
            if (differentFirst)
            {
                //First Header
                entry = PdfTextMap.GetTextFormats(pageSettings, dictionaries, pdfSheet.Worksheet, headerFooter.FirstHeader.LeftAligned, HeaderFooterType.First, HeaderFooterAlignment.Left, HeaderFooterSection.Header);
                if (entry != null)
                {
                    entry.Content.ContentAligmnet = PdfTextMap.GetAlignmentData(entry);
                    PdfHeaderFooterEntries.Add(entry);
                }
                entry = PdfTextMap.GetTextFormats(pageSettings, dictionaries, pdfSheet.Worksheet, headerFooter.FirstHeader.Centered, HeaderFooterType.First, HeaderFooterAlignment.Center, HeaderFooterSection.Header);
                if (entry != null)
                {
                    entry.Content.ContentAligmnet = PdfTextMap.GetAlignmentData(entry);
                    PdfHeaderFooterEntries.Add(entry);
                }
                entry = PdfTextMap.GetTextFormats(pageSettings, dictionaries, pdfSheet.Worksheet, headerFooter.FirstHeader.RightAligned, HeaderFooterType.First, HeaderFooterAlignment.Right, HeaderFooterSection.Header);
                if (entry != null)
                {
                    entry.Content.ContentAligmnet = PdfTextMap.GetAlignmentData(entry);
                    PdfHeaderFooterEntries.Add(entry);
                }
                //First Footer
                entry = PdfTextMap.GetTextFormats(pageSettings, dictionaries, pdfSheet.Worksheet, headerFooter.FirstFooter.LeftAligned, HeaderFooterType.First, HeaderFooterAlignment.Left, HeaderFooterSection.Footer);
                if (entry != null)
                {
                    entry.Content.ContentAligmnet = PdfTextMap.GetAlignmentData(entry);
                    PdfHeaderFooterEntries.Add(entry);
                }
                entry = PdfTextMap.GetTextFormats(pageSettings, dictionaries, pdfSheet.Worksheet, headerFooter.FirstFooter.Centered, HeaderFooterType.First, HeaderFooterAlignment.Center, HeaderFooterSection.Footer);
                if (entry != null)
                {
                    entry.Content.ContentAligmnet = PdfTextMap.GetAlignmentData(entry);
                    PdfHeaderFooterEntries.Add(entry);
                }
                entry = PdfTextMap.GetTextFormats(pageSettings, dictionaries, pdfSheet.Worksheet, headerFooter.FirstFooter.RightAligned, HeaderFooterType.First, HeaderFooterAlignment.Right, HeaderFooterSection.Footer);
                if (entry != null)
                {
                    entry.Content.ContentAligmnet = PdfTextMap.GetAlignmentData(entry);
                    PdfHeaderFooterEntries.Add(entry);
                }
            }
            if (differentOddEven)
            {
                //Even Header
                entry = PdfTextMap.GetTextFormats(pageSettings, dictionaries, pdfSheet.Worksheet, headerFooter.EvenHeader.LeftAligned, HeaderFooterType.Even, HeaderFooterAlignment.Left, HeaderFooterSection.Header);
                if (entry != null)
                {
                    entry.Content.ContentAligmnet = PdfTextMap.GetAlignmentData(entry);
                    PdfHeaderFooterEntries.Add(entry);
                }
                entry = PdfTextMap.GetTextFormats(pageSettings, dictionaries, pdfSheet.Worksheet, headerFooter.EvenHeader.Centered, HeaderFooterType.Even, HeaderFooterAlignment.Center, HeaderFooterSection.Header);
                if (entry != null)
                {
                    entry.Content.ContentAligmnet = PdfTextMap.GetAlignmentData(entry);
                    PdfHeaderFooterEntries.Add(entry);
                }
                entry = PdfTextMap.GetTextFormats(pageSettings, dictionaries, pdfSheet.Worksheet, headerFooter.EvenHeader.RightAligned, HeaderFooterType.Even, HeaderFooterAlignment.Right, HeaderFooterSection.Header);
                if (entry != null)
                {
                    entry.Content.ContentAligmnet = PdfTextMap.GetAlignmentData(entry);
                    PdfHeaderFooterEntries.Add(entry);
                }
                //Even Footer
                entry = PdfTextMap.GetTextFormats(pageSettings, dictionaries, pdfSheet.Worksheet, headerFooter.EvenFooter.LeftAligned, HeaderFooterType.Even, HeaderFooterAlignment.Left, HeaderFooterSection.Footer);
                if (entry != null)
                {
                    entry.Content.ContentAligmnet = PdfTextMap.GetAlignmentData(entry);
                    PdfHeaderFooterEntries.Add(entry);
                }
                entry = PdfTextMap.GetTextFormats(pageSettings, dictionaries, pdfSheet.Worksheet, headerFooter.EvenFooter.Centered, HeaderFooterType.Even, HeaderFooterAlignment.Center, HeaderFooterSection.Footer);
                if (entry != null)
                {
                    entry.Content.ContentAligmnet = PdfTextMap.GetAlignmentData(entry);
                    PdfHeaderFooterEntries.Add(entry);
                }
                entry = PdfTextMap.GetTextFormats(pageSettings, dictionaries, pdfSheet.Worksheet, headerFooter.EvenFooter.RightAligned, HeaderFooterType.Even, HeaderFooterAlignment.Right, HeaderFooterSection.Footer);
                if (entry != null)
                {
                    entry.Content.ContentAligmnet = PdfTextMap.GetAlignmentData(entry);
                    PdfHeaderFooterEntries.Add(entry);
                }
            }
            //Odd Header
            entry = PdfTextMap.GetTextFormats(pageSettings, dictionaries, pdfSheet.Worksheet, headerFooter.OddHeader.LeftAligned, HeaderFooterType.Odd, HeaderFooterAlignment.Left, HeaderFooterSection.Header);
            if (entry != null)
            {
                entry.Content.ContentAligmnet = PdfTextMap.GetAlignmentData(entry);
                PdfHeaderFooterEntries.Add(entry);
            }
            entry = PdfTextMap.GetTextFormats(pageSettings, dictionaries, pdfSheet.Worksheet, headerFooter.OddHeader.Centered, HeaderFooterType.Odd, HeaderFooterAlignment.Center, HeaderFooterSection.Header);
            if (entry != null)
            {
                entry.Content.ContentAligmnet = PdfTextMap.GetAlignmentData(entry);
                PdfHeaderFooterEntries.Add(entry);
            }
            entry = PdfTextMap.GetTextFormats(pageSettings, dictionaries, pdfSheet.Worksheet, headerFooter.OddHeader.RightAligned, HeaderFooterType.Odd, HeaderFooterAlignment.Right, HeaderFooterSection.Header);
            if (entry != null)
            {
                entry.Content.ContentAligmnet = PdfTextMap.GetAlignmentData(entry);
                PdfHeaderFooterEntries.Add(entry);
            }
            //Odd Footer
            entry = PdfTextMap.GetTextFormats(pageSettings, dictionaries, pdfSheet.Worksheet, headerFooter.OddFooter.LeftAligned, HeaderFooterType.Odd, HeaderFooterAlignment.Left, HeaderFooterSection.Footer);
            if (entry != null)
            {
                entry.Content.ContentAligmnet = PdfTextMap.GetAlignmentData(entry);
                PdfHeaderFooterEntries.Add(entry);
            }
            entry = PdfTextMap.GetTextFormats(pageSettings, dictionaries, pdfSheet.Worksheet, headerFooter.OddFooter.Centered, HeaderFooterType.Odd, HeaderFooterAlignment.Center, HeaderFooterSection.Footer);
            if (entry != null)
            {
                entry.Content.ContentAligmnet = PdfTextMap.GetAlignmentData(entry);
                PdfHeaderFooterEntries.Add(entry);
            }
            entry = PdfTextMap.GetTextFormats(pageSettings, dictionaries, pdfSheet.Worksheet, headerFooter.OddFooter.RightAligned, HeaderFooterType.Odd, HeaderFooterAlignment.Right, HeaderFooterSection.Footer);
            if (entry != null)
            {
                entry.Content.ContentAligmnet = PdfTextMap.GetAlignmentData(entry);
                PdfHeaderFooterEntries.Add(entry);
            }
            if (headerFooter?.Pictures != null)
            {
                foreach (ExcelVmlDrawingPicture picture in headerFooter.Pictures)
                {
                    var bytes = picture?.Image?.ImageBytes;
                    if (bytes == null) continue;
                    if (!TryDecodeSlot(picture.Id, out var type, out var section, out var alignment)) continue;
                    PdfHeaderFooterImages.Add(new PdfHeaderFooterImage(bytes, picture.Width, picture.Height, type, alignment, section));
                }
            }
        }

        public PdfHeaderFooterImage GetImage(HeaderFooterType type, HeaderFooterSection section, HeaderFooterAlignment alignment)
        {
            return PdfHeaderFooterImages.FirstOrDefault(e =>
                e.PageType == type && e.Section == section && e.Alignment == alignment);
        }

        // Decode a header/footer picture Id (e.g. "LH", "CFEVEN", "RHFIRST") into its slot.
        // Layout: [alignment L/C/R][section H/F][optional variant EVEN|FIRST]; no variant = Odd.
        private static bool TryDecodeSlot(string id, out HeaderFooterType type, out HeaderFooterSection section, out HeaderFooterAlignment alignment)
        {
            type = HeaderFooterType.Odd;
            section = HeaderFooterSection.Header;
            alignment = HeaderFooterAlignment.Left;
            if (string.IsNullOrEmpty(id) || id.Length < 2) return false;

            switch (id[0])
            {
                case 'L': alignment = HeaderFooterAlignment.Left; break;
                case 'C': alignment = HeaderFooterAlignment.Center; break;
                case 'R': alignment = HeaderFooterAlignment.Right; break;
                default: return false;
            }

            string code = id.Substring(1);
            if (code[0] == 'H') section = HeaderFooterSection.Header;
            else if (code[0] == 'F') section = HeaderFooterSection.Footer;
            else return false;

            string variant = code.Substring(1);
            if (variant.Length == 0) type = HeaderFooterType.Odd;
            else if (variant == "EVEN") type = HeaderFooterType.Even;
            else if (variant == "FIRST") type = HeaderFooterType.First;
            else return false;

            return true;
        }

        public PdfHeaderFooter Get(HeaderFooterType type, HeaderFooterSection section, HeaderFooterAlignment alignment)
        {
            return PdfHeaderFooterEntries.FirstOrDefault(e =>
                e.PageType == type &&
                e.Section == section &&
                e.Alignment == alignment);
        }

        public HeaderFooterType GetPageType(int physicalPageIndex)
        {
            if (physicalPageIndex == 1 && HasFirstPage) return HeaderFooterType.First;
            if (physicalPageIndex % 2 == 0 && HasOddEvenPages) return HeaderFooterType.Even;
            return HeaderFooterType.Odd;
        }
    }
}
