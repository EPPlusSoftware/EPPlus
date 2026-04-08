using EPPlus.Export.Pdf.PdfSettings;
using EPPlus.Graphics;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlus.Export.Pdf.PdfCatalog
{
    internal struct Page
    {
        public int FromRow;
        public int FromColumn;
        public int ToRow;
        public int ToColumn;

        public bool HasPrintTitle;
    }

    internal struct Pages
    {
        public Page[] Page;
        public int Width;
        public int Height;
        public int Count
        {
            get { return Width * Height; }
        }
    }

    internal class PdfLayout
    {
        private const double rowHeadingWith1CharWidth = 23.25d;

        public static Transform GetLayout(PdfPageSettings pageSettings, PdfWorksheet[] pdfSheets)
        {
            List<Pages> PagesCollection = new List<Pages>();
            //calculate number of pages
            foreach (var pdfSheet in pdfSheets)
            {
                foreach (var range in pdfSheet.Ranges)
                {
                    var pages = GetNumberOfPages(pageSettings, pdfSheet, range);
                    AssignRangeToPages(pageSettings, range, pages);
                }
                //add together all pages, assign page number/total page numbers
                //add cells to each page first as an array for gridlines then as transforms
                //create gridlines
                //shape headerfooter text again if it contains page numbers/total pages number.
            }
            return null;
        }

        internal static Pages GetNumberOfPages(PdfPageSettings pageSettings, PdfWorksheet pdfSheet,  PdfRange range)
        {
            //calculte pages needed for this range, add int col headings for width, row headings for height. THis is where we also add print headings later on. Autofit on row here too later on.
            var xPages = (int)Math.Max(1, Math.Ceiling(range.TotalWidth / pageSettings.ContentBounds.Width));
            var yPages = (int)Math.Max(1, Math.Ceiling(range.TotalHeight / pageSettings.ContentBounds.Height));

            if (pageSettings.ShowHeadings)
            {
                int prev = 0;
                do
                {
                    prev = xPages;
                    range.AdditionalWidth = xPages * ((rowHeadingWith1CharWidth - pdfSheet.ZeroCharWidth) + (Math.Abs(pdfSheet.ToRow).ToString().Length * pdfSheet.ZeroCharWidth));
                    xPages = (int)Math.Max(1, Math.Ceiling((range.TotalWidth + range.AdditionalWidth) / pageSettings.ContentBounds.Width));
                } while (prev != xPages);
                do
                {
                    prev = yPages;
                    range.AdditionalHeight = yPages * pdfSheet.Worksheet.DefaultRowHeight;
                    yPages = (int)Math.Max(1, Math.Ceiling((range.TotalHeight + range.AdditionalHeight) / pageSettings.ContentBounds.Height));
                } while (prev != yPages);
            }
            for (int i = range.Range._fromCol; i <= range.Range._toCol; i++)
            {
                if (pdfSheet.Worksheet.Column(i).PageBreak)
                    xPages++;
            }
            for (int i = range.Range._fromRow; i <= range.Range._toRow; i++)
            {
                if (pdfSheet.Worksheet.Row(i).PageBreak)
                    yPages++;
            }
            //might need to place ranges into pages before checking out titles since what page titles start showing on might diff. Maybe page is
            //if (HasPrintTitles Row)
            //if (HasPrintTitles Column)

            Pages p;
            p.Width = xPages;
            p.Height = yPages;
            p.Page = null;
            return p;
        }

        internal static Pages AssignRangeToPages(PdfPageSettings pageSettings, PdfRange range, Pages pdfPages)
        {
            var pages = pdfPages;
            var worksheet = range.Range.Worksheet;
            pages.Page = new Page[pages.Count];
            int col = 0, row = 0;
            int fromCol = range.Map.FromColumn;
            int fromRow = range.Map.FromRow;
            var addedWidth = range.AdditionalWidth / pages.Width;
            var addedHeight = range.AdditionalHeight / pages.Height;
            for (int i = 0; i < pages.Count; i++)
            {
                var page = pages.Page[i];
                double width = 0d, height = 0d;
                while (col < range.ColWidths.Count)
                {
                    int actualCol = range.Range._fromCol + col;
                    if (width + range.ColWidths[col] + addedWidth >= pageSettings.ContentBounds.Width)
                    {
                        page.FromColumn = fromCol;
                        page.ToColumn = fromCol + col;
                        fromCol += col;
                        break;
                    }
                    width += range.ColWidths[col];
                    col++;
                    if (worksheet.Column(actualCol).PageBreak)
                    {
                        page.FromColumn = fromCol;
                        page.ToColumn = fromCol + col - 1;
                        fromCol += col;
                        width = 0d;
                        break;
                    }
                }
                while (row < range.RowHeights.Count)
                {
                    int actualRow = range.Range._fromRow + row;
                    if (height + range.RowHeights[row] + addedHeight >= pageSettings.ContentBounds.Height)
                    {
                        page.FromRow = fromRow;
                        page.ToRow = fromRow + row;
                        fromRow += row;
                        break;
                    }
                    height += range.RowHeights[row];
                    row++;
                    if (worksheet.Row(actualRow).PageBreak)
                    {
                        page.FromColumn = fromCol;
                        page.ToColumn = fromCol + col - 1;
                        fromCol += col;
                        width = 0d;
                        break;
                    }
                }
                if (i == pages.Count - 1)
                {
                    page.FromColumn = fromCol;
                    page.ToColumn = fromCol + col;
                    page.FromRow = fromRow;
                    page.ToRow = fromRow + row;
                }
                pages.Page[i] = page;
            }
            pdfPages = pages;
            return pdfPages;
        }
    }
}
