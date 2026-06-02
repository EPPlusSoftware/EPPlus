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
using EPPlus.Export.Pdf.PdfLayout;
using EPPlus.Export.Pdf.PdfObjects;
using EPPlus.Export.Pdf.PdfResources;
using EPPlus.Export.Pdf.PdfSettings;
using EPPlus.Fonts.OpenType;
using EPPlus.Graphics;
using OfficeOpenXml;
using OfficeOpenXml.Interfaces.Fonts;
using OfficeOpenXml.Style;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using static OfficeOpenXml.Drawing.OleObject.Structures.OleObjectDataStructures;

namespace EPPlus.Export.Pdf
{
    /// <summary>
    /// Class for exporting to PDF format.
    /// </summary>
    public class ExcelPdf
    {
        internal List<ExcelWorksheet> _workheets = new List<ExcelWorksheet>();
        internal ExcelRangeBase _range;
        private PdfPageSettings _pageSettings;
        internal List<PdfObject> _document = new List<PdfObject>();
        internal string header = "%PDF-1.7\n";
        internal PdfDictionaries _dictionaries = new PdfDictionaries();
        private string _debugString;

        public ExcelPdf()
        {
        }

        /// <summary>
        /// Create a PDF Document from the worksheet and settings.
        /// </summary>
        /// <param name="worksheet">The worksheet to convert to PDF Document</param>
        /// <param name="pageSettings">The settings object</param>
        public ExcelPdf(ExcelWorksheet worksheet, PdfPageSettings pageSettings = null)
        {
            _workheets.Add(worksheet);
            _pageSettings = pageSettings == null ? new PdfPageSettings() : pageSettings;
            _pageSettings.defaultFontName = worksheet.Workbook.ThemeManager.CurrentTheme.FontScheme.MinorFont[0].Typeface;
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
                if (_dictionaries.Patterns.ContainsKey(patternName))
                {
                    return _dictionaries.Patterns[patternName].Label;
                }
            }
            return null;
        }

        //Add Fonts //Need to update this method a bit. We should check for all default fonts and not only courier new? Also need to check if we are allowed to embedd the font.
        internal void AddFontData()
        {
            if (_pageSettings.EmbeddFonts)
            {
                foreach (var font in _dictionaries.Fonts)
                {
                    //font.Value.CreateGidsAndCharMaps();
                    var cidSet = font.Value.GetCidSet(_document.Count + 1);
                    if (cidSet != null) _document.Add(cidSet);
                    _document.Add(font.Value.GetEmbeddedFontStreamObject(_document.Count + 1));
                    _document.Add(font.Value.GetFontDescriptorObject(_document.Count + 1));
                    _document.Add(font.Value.GetCIDFontObject(_document.Count + 1));
                    _document.Add(font.Value.GetUnicodeCmapObject(_document.Count + 1));
                    _document.Add(font.Value.GetType0FontDictObject(_document.Count + 1));
                    font.Value.GetFontObject(_document.Count);
                }
            }
            else
            {
                foreach (var font in _dictionaries.Fonts)
                {
                    _document.Add(font.Value.GetFontDescriptorObject(_document.Count + 1));
                    _document.Add(font.Value.GetWidthsObject(_document.Count + 1));
                    _document.Add(font.Value.GetFontObject(_document.Count + 1));
                }
            }
        }

        //Add Patterns
        internal void AddPatternData()
        {
            foreach (var pattern in _dictionaries.Patterns)
            {
                _document.Add(pattern.Value.GetPatternObject(_document.Count + 1));
            }
        }

        //Add Shadings and accompanying pattern
        internal void AddShadingsData()
        {
            foreach (var shading in _dictionaries.Shadings)
            {
                _document.Add(shading.Value.GetShadingObject(_document.Count + 1));
                _document.Add(shading.Value.GetShadingPatternObject(_document.Count + 1, _document.Count));
                int label = _dictionaries.Patterns.Last().Value.labelNumber + 1;
                var pr = new PdfPatternResource(label, shading.Value.CellFillData);
                pr.objectNumber = _document.Count;
                _dictionaries.Patterns.Add(shading.Value.CellFillData.id, pr);
            }
        }

        //Create Page
        private PdfPage AddPage(int pagesObjectNumber, List<int> contentObjectNumbers, PdfPageSettings settings)
        {
            var page = new PdfPage(_document.Count + 1, pagesObjectNumber, contentObjectNumbers, settings.PageSize, _dictionaries);
            _document.Add(page);
            return page;
        }

        //Create Pages
        private PdfPages AddPages()
        {
            var pages = new PdfPages(_document.Count + 1, new List<int> { });
            _document.Add(pages);
            return pages;
        }

        //Create Catalog
        private PdfObjects.PdfCatalog AddCatalog(int pagesObjectNumber)
        {
            var catalog = new PdfObjects.PdfCatalog(_document.Count + 1, pagesObjectNumber);
            _document.Add(catalog);
            return catalog;
        }

        //Create Content
        private void AddContent(Transform pageLayout, PdfPage page)
        {
            //var cells = pageLayout.ChildObjects.Where(t => t is PdfCellLayout || t is PdfCellContentLayout || t is PdfCellBorderLayout).GroupBy(t => t.Name);
            var cells = pageLayout.ChildObjects.Where(t => (t is PdfCellLayout || t is PdfCellContentLayout || t is PdfCellBorderLayout) && !(t is PdfCellContentLayout cc && cc.IsHeaderFooter)).GroupBy(t => t.Name);
            var headerFooterLayouts = pageLayout.ChildObjects.OfType<PdfCellContentLayout>().Where(t => t.IsHeaderFooter);
            var contentStream = new PdfContentStream(_document.Count + 1);
            contentStream.AddCommand($"% {pageLayout.Name} start");
            //Add clipping rectangle around page content.
            contentStream.AddCommand("q");
            contentStream.AddMarginClipping((PdfPageLayout)pageLayout);
            if (_pageSettings.ShowGridLines)
            {
                contentStream.AddInnerGridLines(pageLayout);
            }
            foreach (var cell in cells)
            {
                foreach (var cellPart in cell)
                {
                    contentStream.AddCommand($"% CELL : {cellPart.Name}");
                    switch (cellPart)
                    {
                        case PdfCellLayout layout:
                            contentStream.AddCellLayout(layout, GetPatternLabel(layout));
                            break;
                        case PdfCellContentLayout contentLayout:
                            contentStream.AddCellContentLayout(contentLayout, _dictionaries, _pageSettings);
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
            if (_pageSettings.ShowGridLines)
            {
                contentStream.AddOuterGridBorder(pageLayout);
            }
            //Add header and footer.
            foreach (var hf in headerFooterLayouts)
            {
                contentStream.AddCellContentLayout(hf, _dictionaries, _pageSettings);
            }
            _document.Add(contentStream);
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
                contentStream.AddCellContentLayout(headerFooterLayout, _dictionaries, _pageSettings);
            }
        }

        //Add Info
        private PdfInfoObject AddInfoObject(string workBookName = "")
        {
            var info = new PdfInfoObject(_document.Count + 1, workBookName);
            _document.Add(info);
            return info;
        }

        internal void CreatePdf(PdfPageSettings pageSettings, PdfDictionaries dictionaries, Transform layout, string fileName)
        {
            using (var fs = new FileStream(fileName, FileMode.Create, FileAccess.Write))
            {
                CreatePdf(pageSettings, dictionaries, layout, fs);
            }

            if (_pageSettings.Debug && _pageSettings.PrintAsText)
            {
                WriteDebugText(fileName);
            }
        }

        internal void CreatePdf(PdfPageSettings pageSettings, PdfDictionaries dictionaries, Transform layout, Stream stream)
        {
            _pageSettings = pageSettings;
            _dictionaries = dictionaries;

            var catalog = AddCatalog(2);
            //Create Pages
            var pagesLayout = layout.ChildObjects[0];
            var pages = AddPages();
            //Create Fonts
            AddFontData();
            //Create Patterns
            AddPatternData();
            //Create Shadings
            AddShadingsData();
            //Create Page and Content
            for (int i = 0; i < layout.ChildObjects.Count; i++)
            {
                var pageLayout = layout.ChildObjects[i];
                var page = AddPage(2, new List<int>(), _pageSettings);
                AddContent(pageLayout, page);
                pages.pageObjectNumbers.Add(page.objectNumber);
            }
            var info = AddInfoObject();

            WriteDocumentToStream(stream, catalog, info);
        }

        /// <summary>
        /// Create the pdf from the supplied worksheet.
        /// </summary>
        /// <param name="Filename">The file name</param>
        public void CreatePdf(string Filename)
        {
            using (var fs = new FileStream(Filename, FileMode.Create, FileAccess.Write))
            {
                CreatePdf(fs);
            }

            if (_pageSettings.Debug && _pageSettings.PrintAsText)
            {
                WriteDebugText(Filename);
            }
        }

        /// <summary>
        /// Create the pdf from the supplied worksheet and write it to a stream.
        /// </summary>
        /// <param name="stream">The stream to write the pdf to. The stream will not be closed.</param>
        public void CreatePdf(Stream stream)
        {
            //Create Catalog
            var catalogLayout = new PdfCatalogLayout(_workheets[0], _pageSettings, _dictionaries);
            var catalog = AddCatalog(2);
            //Create Pages
            var pagesLayout = catalogLayout.ChildObjects[0];
            var pages = AddPages();
            //Create Fonts
            AddFontData();
            //Create Patterns
            AddPatternData();
            //Create Shadings
            AddShadingsData();
            //Create Page and Content
            for (int i = 0; i < pagesLayout.ChildObjects.Count; i++)
            {
                var pageLayout = pagesLayout.ChildObjects[i];
                var page = AddPage(2, new List<int>(), _pageSettings);
                AddContent(pageLayout, page);
                pages.pageObjectNumbers.Add(page.objectNumber);
            }
            var info = AddInfoObject(_workheets[0].Workbook._package.File.Name);

            WriteDocumentToStream(stream, catalog, info);
        }

        //Write the document and cross-ref/trailer to the supplied stream.
        //The stream is not closed; the caller owns it.
        private void WriteDocumentToStream(Stream stream, PdfObjects.PdfCatalog catalog, PdfInfoObject info)
        {
            _debugString = "";
            PdfCrossRefTable crossRefTable = new PdfCrossRefTable();

            //Use a BinaryWriter without disposing it, so the underlying stream stays open for the caller.
            //BinaryWriter does not own the stream when we don't dispose it; we just flush at the end.
            var bw = new BinaryWriter(stream, Encoding.ASCII);
            try
            {
                //Write header
                bw.Write(Encoding.ASCII.GetBytes(header));
                _debugString += header;
                //Write body
                foreach (var pdfobj in _document)
                {
                    crossRefTable.AddPosition(stream.Position);
                    pdfobj.ToPdfBytes(bw);
                    _debugString += pdfobj.ToPdfString();
                }
                //Write CrossReference
                crossRefTable.Write(bw, stream.Position, _document.Count);
                _debugString += crossRefTable.WriteString(_document.Count);
                // Write trailer
                PdfTrailer.Write(bw, _document.Count, catalog.objectNumber, info.objectNumber, crossRefTable.StartPosition);
                _debugString += PdfTrailer.WriteString(_document.Count, catalog.objectNumber, info.objectNumber, crossRefTable.StartPosition);
            }
            finally
            {
                bw.Flush();
            }
        }

        private void WriteDebugText(string fileName)
        {
            using (var fs = new FileStream(fileName + ".txt", FileMode.Create, FileAccess.Write))
            {
                using (var wr = new StreamWriter(fs))
                {
                    wr.Write(_debugString);
                }
            }
        }
    }
}