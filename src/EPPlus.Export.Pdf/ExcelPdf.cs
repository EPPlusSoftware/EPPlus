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
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using OfficeOpenXml;
using EPPlus.Graphics;
using EPPlus.Export.Pdf.PdfLayout;
using EPPlus.Export.Pdf.PdfObjects;
using EPPlus.Export.Pdf.PdfResources;
using EPPlus.Export.Pdf.PdfSettings;
using OfficeOpenXml.Style;
using EPPlus.Fonts.OpenType;
using OfficeOpenXml.Interfaces.Fonts;

namespace EPPlus.Export.Pdf
{
    /// <summary>
    /// Class for exporting to PDF format.
    /// </summary>
    public class ExcelPdf
    {
        internal List<ExcelWorksheet> _workheets = new List<ExcelWorksheet>();
        internal ExcelRangeBase _range;
        private PdfPageSettings PageSettings;
        internal List<PdfObject> Document = new List<PdfObject>();
        internal string header = "%PDF-1.7\n";
        internal PdfDictionaries Dictionaries = new PdfDictionaries();

        /// <summary>
        /// Create a PDF Document from the worksheet and settings.
        /// </summary>
        /// <param name="worksheet">The worksheet to convert to PDF Document</param>
        /// <param name="pageSettings">The settings object</param>
        public ExcelPdf(ExcelWorksheet worksheet, PdfPageSettings pageSettings = null)
        {
            _workheets.Add(worksheet);
            PageSettings = pageSettings == null ? new PdfPageSettings() : pageSettings;
            PageSettings.defaultFontName = worksheet.Workbook.ThemeManager.CurrentTheme.FontScheme.MinorFont[0].Typeface;
        }

        /// <summary>
        /// Create a PDF Document from the selected worksheets and settings. NOT IMPLEMENTED
        /// </summary>
        /// <param name="worksheet">The worksheets to convert to PDF Document</param>
        /// <param name="pageSettings">The Settings object</param>
        public ExcelPdf(ExcelWorksheet[] worksheet, PdfPageSettings pageSettings = null)
        {
            //_ws = worksheet[0];
            //defaultFontName = worksheet[0].Workbook.ThemeManager.CurrentTheme.FontScheme.MinorFont[0].Typeface;
            //PageSettings = pageSettings == null ? new PdfPageSettings() : pageSettings;
        }

        /// <summary>
        /// Create a PDF Document from the entire worksbook and settings.NOT IMPLEMENTED
        /// </summary>
        /// <param name="workbook">Workbook to convert to PDF Document</param>
        /// <param name="pageSettings">The settings object</param>
        public ExcelPdf(ExcelWorkbook workbook, PdfPageSettings pageSettings = null)
        {
            //_ws = worksheet;
            //defaultFontName = worksheet.Workbook.ThemeManager.CurrentTheme.FontScheme.MinorFont[0].Typeface;
            //PageSettings = pageSettings == null ? new PdfPageSettings() : pageSettings;
        }

        /// <summary>
        /// Create a PDF Document from the selected range and settings. NOT IMPLEMENTED
        /// </summary>
        /// <param name="Range">Range to convert to PDF Document</param>
        /// <param name="pageSettings">The settings object</param>
        public ExcelPdf(ExcelRangeBase Range, PdfPageSettings pageSettings = null)
        {
            //_range = Range;
            //defaultFontName = Range.Worksheet.Workbook.ThemeManager.CurrentTheme.FontScheme.MinorFont[0].Typeface;
            //PageSettings = pageSettings == null ? new PdfPageSettings() : pageSettings;
        }

        //Get the label to use for pattern.
        internal string GetPatternLabel(PdfCellLayout layout)
        {
            if ((layout.CellFillData.PatternStyle != ExcelFillStyle.Solid && layout.CellFillData.PatternStyle != ExcelFillStyle.None) || layout.CellFillData.GradientFillData != null)
            {
                var patternName = layout.CellFillData.id;
                if (Dictionaries.Patterns.ContainsKey(patternName))
                {
                    return Dictionaries.Patterns[patternName].Label;
                }
            }
            return null;
        }

        //Add Fonts //Need to update this method a bit. We should check for all default fonts and not only courier new? Also need to check if we are allowed to embedd the font.
        internal void AddFontData()
        {
            if (PageSettings.EmbeddFonts)
            {
                foreach (var font in Dictionaries.Fonts)
                {
                    //font.Value.CreateGidsAndCharMaps();
                    var CidSet = font.Value.GetCidSet(Document.Count + 1);
                    if (CidSet != null) Document.Add(CidSet);
                    Document.Add(font.Value.GetEmbeddedFontStreamObject(Document.Count + 1));
                    Document.Add(font.Value.GetFontDescriptorObject(Document.Count + 1));
                    Document.Add(font.Value.GetCIDFontObject(Document.Count + 1));
                    Document.Add(font.Value.GetUnicodeCmapObject(Document.Count + 1));
                    Document.Add(font.Value.GetType0FontDictObject(Document.Count + 1));
                    font.Value.GetFontObject(Document.Count);
                }
            }
            else
            {
                foreach (var font in Dictionaries.Fonts)
                {
                    Document.Add(font.Value.GetFontDescriptorObject(Document.Count + 1));
                    Document.Add(font.Value.GetWidthsObject(Document.Count + 1));
                    Document.Add(font.Value.GetFontObject(Document.Count + 1));
                }
            }
        }

        //Add Patterns
        internal void AddPatternData()
        {
            foreach (var pattern in Dictionaries.Patterns)
            {
                Document.Add(pattern.Value.GetPatternObject(Document.Count + 1));
            }
        }

        //Add Shadings and accompanying pattern
        internal void AddShadingsData(PdfDictionaries dictionaries)
        {
            foreach (var shading in Dictionaries.Shadings)
            {
                Document.Add(shading.Value.GetShadingObject(Document.Count + 1));
                Document.Add(shading.Value.GetShadingPatternObject(Document.Count + 1, Document.Count));
                int label = dictionaries.Patterns.Last().Value.labelNumber + 1;
                var pr = new PdfPatternResource(label, shading.Value.CellFillData);
                pr.objectNumber = Document.Count;
                dictionaries.Patterns.Add(shading.Value.CellFillData.id, pr);
            }
        }

        //Create Page
        private PdfPage AddPage(int pagesObjectNumber, List<int> contentObjectNumbers, PdfPageSettings settings)
        {
            var page = new PdfPage(Document.Count + 1, pagesObjectNumber, contentObjectNumbers, settings.PageSize, Dictionaries);
            Document.Add(page);
            return page;
        }

        //Create Pages
        private PdfPages AddPages()
        {
            var pages = new PdfPages(Document.Count + 1, new List<int>{});
            Document.Add(pages);
            return pages;
        }

        //Create Catalog
        private PdfCatalog AddCatalog(int pagesObjectNumber)
        {
            var catalog = new PdfCatalog(Document.Count + 1, pagesObjectNumber);
            Document.Add(catalog);
            return catalog;
        }

        //Create Content
        private void AddContent(Transform pageLayout, PdfPage page)
        {
            var cells = pageLayout.ChildObjects.Where(t => t is PdfCellLayout || t is PdfCellContentLayout || t is PdfCellBorderLayout).GroupBy(t => t.Name);
            var contentStream = new PdfContentStream(Document.Count + 1);
            contentStream.AddCommand($"% {pageLayout.Name} start");
            //Add clipping rectangle around page content.
            contentStream.AddCommand("q");
            //contentStream.AddMarginClipping((PdfPageLayout)pageLayout);
            if (PageSettings.ShowGridLines)
            {
                contentStream.AddInnerGridLines(pageLayout);
            }
            foreach (var cell in cells)
            {
                foreach (var cellPart in cell)
                {
                    switch (cellPart)
                    {
                        case PdfCellLayout layout:
                            contentStream.AddCellLayout(layout, GetPatternLabel(layout));
                            break;
                        case PdfCellContentLayout contentLayout:
                            contentStream.AddCellContentLayout(contentLayout, Dictionaries, PageSettings);
                            break;
                        case PdfCellBorderLayout borderLayout:
                            contentStream.AddBorderLayout(borderLayout);
                            break;
                    }
                }
            }
            //Close the clipping rectangle.
            contentStream.AddCommand("Q");
            contentStream.AddCommand($"% Margin Clip End");
            if (PageSettings.ShowGridLines)
            {
                contentStream.AddOuterGridBorder(pageLayout);
            }
            //Add header and footer.
            AddHeaderFooter(contentStream, pageLayout, page);
            Document.Add(contentStream);
            page.contentObjectNumbers.Add(contentStream.objectNumber);
            contentStream.AddCommand($"% {pageLayout.Name} end");
        }

        //Add Header Footer
        private void AddHeaderFooter(PdfContentStream contentStream, Transform pageLayout, PdfPage page)
        {
            var headerFooter = pageLayout.ChildObjects.Where(t => t is PdfHeaderFooterLayout);
            foreach (var hf in headerFooter)
            {
                var headerFooterLayout = hf as PdfHeaderFooterLayout;
                contentStream.AddCellContentLayout(headerFooterLayout, Dictionaries, PageSettings);
            }
        }

        //Add Info
        private PdfInfoObject AddInfoObject()
        {
            var info = new PdfInfoObject(Document.Count + 1, _workheets[0].Workbook._package.File.Name);
            Document.Add(info);
            return info;
        }

        /// <summary>
        /// Create the pdf from the supplied worksheet.
        /// </summary>
        /// <param name="Filename">The file name</param>
        public void CreatePdf(string Filename)
        {
            //Create Catalog
            var catalogLayout = new PdfCatalogLayout(_workheets[0], PageSettings, Dictionaries);
            var catalog = AddCatalog(2);
            //Create Pages
            var pagesLayout = catalogLayout.ChildObjects[0];
            var pages = AddPages();
            //Create Fonts
            AddFontData();
            //Create Patterns
            AddPatternData();
            //Create Shadings
            AddShadingsData(Dictionaries);
            //Create Page and Content
            for (int i = 0; i < pagesLayout.ChildObjects.Count; i++)
            {
                var pageLayout = pagesLayout.ChildObjects[i];
                var page = AddPage(2, new List<int>(), PageSettings);
                AddContent(pageLayout, page);
                pages.pageObjectNumbers.Add(page.objectNumber);
            }
            var info = AddInfoObject();
            string debugString = "";
            //write to pdf
            PdfCrossRefTable crossRefTable = new PdfCrossRefTable();
            //start wring pdf binary
            using (var fs = new FileStream(Filename, FileMode.Create, FileAccess.Write))
            {
                using (var bw = new BinaryWriter(fs, Encoding.ASCII))
                {
                    //Write header
                    bw.Write(Encoding.ASCII.GetBytes(header));
                    debugString += header;
                    //Write body
                    foreach (var pdfobj in Document)
                    {
                        crossRefTable.AddPosition(fs.Position);
                        pdfobj.ToPdfBytes(bw);
                        debugString += pdfobj.ToPdfString();
                    }
                    //Write CrossReference
                    crossRefTable.Write(bw, fs.Position, Document.Count);
                    debugString += crossRefTable.WriteString(Document.Count);
                    // Write trailer
                    PdfTrailer.Write(bw, Document.Count, catalog.objectNumber, info.objectNumber, crossRefTable.StartPosition);
                    debugString += PdfTrailer.WriteString(Document.Count, catalog.objectNumber, info.objectNumber, crossRefTable.StartPosition);
                }
            }
            //Write pdf as txt for debug.
            if (PageSettings.Debug && PageSettings.PrintAsText)
            {
                using (var fs = new FileStream(Filename + ".txt", FileMode.Create, FileAccess.Write))
                {
                    using ( var wr = new StreamWriter(fs))
                    {
                        wr.Write(debugString);
                    }
                }
            }
        }
    }
}
