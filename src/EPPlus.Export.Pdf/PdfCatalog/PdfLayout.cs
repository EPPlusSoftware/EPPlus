using EPPlus.Export.Pdf.PdfSettings;
using EPPlus.Graphics;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlus.Export.Pdf.PdfCatalog
{
    internal struct Pages
    {
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
            //calculate number of pages
            foreach (var pdfSheet in pdfSheets)
            {
                foreach (var range in pdfSheet.Ranges)
                {
                    //calculte pages needed for this range, add int col headings for width, row headings for height. THis is where we also add print headings later on. Autofit on row here too later on.
                    GetNumberOfPages(pageSettings, pdfSheet, range);

                    //create a temp page that contains the range for said page
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
                    var additionalWidth = xPages * ((rowHeadingWith1CharWidth - pdfSheet.ZeroCharWidth) + (Math.Abs(pdfSheet.ToRow).ToString().Length * pdfSheet.ZeroCharWidth));
                    xPages = (int)Math.Max(1, Math.Ceiling((range.TotalWidth + additionalWidth) / pageSettings.ContentBounds.Width));
                } while (prev != xPages);
                do
                {
                    prev = yPages;
                    var additionalHeight = yPages * pdfSheet.Worksheet.DefaultRowHeight;
                    yPages = (int)Math.Max(1, Math.Ceiling((range.TotalHeight + additionalHeight) / pageSettings.ContentBounds.Height));
                } while (prev != yPages);
            }
            //if (HasPrintTitles Row)
            //if (HasPrintTitles Column)

            Pages p;
            p.Width = xPages;
            p.Height = yPages;
            return p;
        }
    }
}
