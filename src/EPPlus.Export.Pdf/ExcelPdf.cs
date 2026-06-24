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
using EPPlus.Export.Pdf.DocumentObjects;
using EPPlus.Export.Pdf.Enums;
using EPPlus.Export.Pdf.Layout;
using EPPlus.Export.Pdf.Resources;
using EPPlus.Export.Pdf.Settings;
using EPPlus.Graphics;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace EPPlus.Export.Pdf
{
    /// <summary>
    /// Class for exporting to PDF format.
    /// </summary>
    internal class ExcelPdf
    {
        private PdfPageSettings _pageSettings;
        private PdfDictionaries _dictionaries;
        private List<PdfObject> _document = new List<PdfObject>();

        private string _debugString;

        internal static string Header
        {
            get
            {
                return "%PDF-1.7\n";
            }
        }

        //Get the label to use for pattern.
        private string GetPatternLabel(PdfCellLayout layout)
        {
            bool isPattern = (layout.CellFillData.PatternStyle != ExcelFillStyle.Solid && layout.CellFillData.PatternStyle != ExcelFillStyle.None) || layout.CellFillData.GradientFillData != null;
            if (isPattern)
            {
                var patternName = layout.CellFillData.id;
                if (patternName != null && _dictionaries.Patterns.ContainsKey(patternName))
                {
                    return _dictionaries.Patterns[patternName].Label;
                }
            }
            return null;
        }

        //Add Fonts //Need to update this method a bit. We should check for all default fonts and not only courier new? Also need to check if we are allowed to embedd the font.
        private void AddFontData()
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
        private void AddPatternData()
        {
            foreach (var pattern in _dictionaries.Patterns)
            {
                _document.Add(pattern.Value.GetPatternObject(_document.Count + 1));
            }
        }

        //Add Shadings and accompanying pattern
        private void AddShadingsData()
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
        private PdfCatalog AddCatalog(int pagesObjectNumber)
        {
            var catalog = new PdfCatalog(_document.Count + 1, pagesObjectNumber);
            _document.Add(catalog);
            return catalog;
        }

        //Create Content
        private void AddContent(Transform pageLayout, PdfPage page)
        {
            var cells = pageLayout.ChildObjects.Where(t =>
                                                (t is PdfCellLayout || t is PdfCellContentLayout || t is PdfCellBorderLayout) &&
                                                !(t is PdfCellLayout cc && (cc.IsHeading || cc.IsPrintTitle)) &&
                                                !(t is PdfCellContentLayout ccl && (ccl.IsHeaderFooter || ccl.IsHeading || ccl.IsPrintTitle)) &&
                                                !(t is PdfCellBorderLayout cbl && cbl.IsPrintTitle)).GroupBy(t => t.Name);


            var headerFooterLayouts = pageLayout.ChildObjects.OfType<PdfCellContentLayout>().Where(t => t.IsHeaderFooter);
            var headingLayouts = pageLayout.ChildObjects.Where(t => (t is PdfCellLayout cl && cl.IsHeading) || (t is PdfCellContentLayout ccl && ccl.IsHeading));
            var printTitleLayouts = pageLayout.ChildObjects.Where(t => (t is PdfCellLayout pl && pl.IsPrintTitle) || (t is PdfCellContentLayout pcl && pcl.IsPrintTitle) || (t is PdfCellBorderLayout pbl && pbl.IsPrintTitle));
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
            // Heading cells render outside the clip — no merged-cell content can obscure them.
            foreach (var heading in headingLayouts)
            {
                contentStream.AddCommand($"% HEADING : {heading.Name}");
                switch (heading)
                {
                    case PdfCellLayout layout:
                        contentStream.AddCellLayout(layout, GetPatternLabel(layout)); break;
                    case PdfCellContentLayout contentLayout:
                        contentStream.AddCellContentLayout(contentLayout, _dictionaries, _pageSettings); break;
                    case PdfCellBorderLayout borderLayout:
                        contentStream.AddBorderLayout(borderLayout); break;
                }
            }
            if (_pageSettings.ShowGridLines || _pageSettings.ShowHeadings)
            {
                contentStream.AddOuterGridBorder(pageLayout);
                contentStream.AddPrintTitleGridLines(pageLayout);
            }
            //Add header and footer.
            foreach (var hf in headerFooterLayouts)
            {
                contentStream.AddCellContentLayout(hf, _dictionaries, _pageSettings);
            }
            foreach (var titleCell in printTitleLayouts)
            {
                contentStream.AddCommand($"% PRINT TITLE : {titleCell.Name}");
                switch (titleCell)
                {
                    case PdfCellLayout layout:
                        contentStream.AddCellLayout(layout, GetPatternLabel(layout)); break;
                    case PdfCellContentLayout contentLayout:
                        contentStream.AddCellContentLayout(contentLayout, _dictionaries, _pageSettings); break;
                    case PdfCellBorderLayout borderLayout:
                        contentStream.AddBorderLayout(borderLayout); break;
                }
            }
            _document.Add(contentStream);
            page.contentObjectNumbers.Add(contentStream.objectNumber);
            contentStream.AddCommand($"% {pageLayout.Name} end");
        }

        //Add Header Footer
        //private void AddHeaderFooter(PdfContentStream contentStream, Transform pageLayout, PdfPage page)
        //{
        //    var headerFooter = pageLayout.ChildObjects.Where(t => t is PdfHeaderFooterLayout);
        //    foreach (var hf in headerFooter)
        //    {
        //        var headerFooterLayout = hf as PdfHeaderFooterLayout;
        //        contentStream.AddCellContentLayout(headerFooterLayout, _dictionaries, _pageSettings);
        //    }
        //}

        //Add Info
        private PdfInfoObject AddInfoObject(string workBookName = "")
        {
            var info = new PdfInfoObject(_document.Count + 1, workBookName);
            _document.Add(info);
            return info;
        }

        internal void CreatePdf(PdfPageSettings pageSettings, PdfDictionaries dictionaries, Transform layout, string fileName)
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
            string debugString = "";
            //write to pdf
            PdfCrossRefTable crossRefTable = new PdfCrossRefTable();
            //start wring pdf binary
            using (var fs = new FileStream(fileName, FileMode.Create, FileAccess.Write))
            {
                using (var bw = new BinaryWriter(fs, Encoding.ASCII))
                {
                    //Write header
                    bw.Write(Encoding.ASCII.GetBytes(Header));
                    debugString += Header;
                    //Write body
                    foreach (var pdfobj in _document)
                    {
                        crossRefTable.AddPosition(fs.Position);
                        pdfobj.ToPdfBytes(bw);
                        debugString += pdfobj.ToPdfString();
                    }
                    //Write CrossReference
                    crossRefTable.Write(bw, fs.Position, _document.Count);
                    debugString += crossRefTable.WriteString(_document.Count);
                    // Write trailer
                    PdfTrailer.Write(bw, _document.Count, catalog.objectNumber, info.objectNumber, crossRefTable.StartPosition);
                    debugString += PdfTrailer.WriteString(_document.Count, catalog.objectNumber, info.objectNumber, crossRefTable.StartPosition);
                }
            }
            //Write pdf as txt for debug.
            if (_pageSettings.Debug && _pageSettings.PrintAsText)
            {
                using (var fs = new FileStream(fileName + ".txt", FileMode.Create, FileAccess.Write))
                {
                    using (var wr = new StreamWriter(fs))
                    {
                        wr.Write(debugString);
                    }
                }
            }
        }


        //internal void CreatePdf(PdfPageSettings pageSettings, PdfDictionaries dictionaries, Transform layout, string fileName)
        //{
        //    using (var fs = new FileStream(fileName, FileMode.Create, FileAccess.Write))
        //    {
        //        CreatePdf(pageSettings, dictionaries, layout, fs);
        //    }

        //    if (_pageSettings.Debug && _pageSettings.PrintAsText)
        //    {
        //        WriteDebugText(fileName);
        //    }
        //}

        //internal void CreatePdf(PdfPageSettings pageSettings, PdfDictionaries dictionaries, Transform layout, Stream stream)
        //{
        //    _pageSettings = pageSettings;
        //    _dictionaries = dictionaries;

        //    var catalog = AddCatalog(2);
        //    //Create Pages
        //    var pagesLayout = layout.ChildObjects[0];
        //    var pages = AddPages();
        //    //Create Fonts
        //    AddFontData();
        //    //Create Patterns
        //    AddPatternData();
        //    //Create Shadings
        //    AddShadingsData();
        //    //Create Page and Content
        //    for (int i = 0; i < layout.ChildObjects.Count; i++)
        //    {
        //        var pageLayout = layout.ChildObjects[i];
        //        var page = AddPage(2, new List<int>(), _pageSettings);
        //        AddContent(pageLayout, page);
        //        pages.pageObjectNumbers.Add(page.objectNumber);
        //    }
        //    var info = AddInfoObject();

        //    WriteDocumentToStream(stream, catalog, info);
        //}

        ///// <summary>
        ///// Create the pdf from the supplied worksheet.
        ///// </summary>
        ///// <param name="Filename">The file name</param>
        //internal void CreatePdf(string Filename)
        //{
        //    using (var fs = new FileStream(Filename, FileMode.Create, FileAccess.Write))
        //    {
        //        CreatePdf(fs);
        //    }

        //    if (_pageSettings.Debug && _pageSettings.PrintAsText)
        //    {
        //        WriteDebugText(Filename);
        //    }
        //}

        /////// <summary>
        /////// Create the pdf from the supplied worksheet and write it to a stream.
        /////// </summary>
        /////// <param name="stream">The stream to write the pdf to. The stream will not be closed.</param>
        ////internal void CreatePdf(Stream stream)
        ////{
        ////    //Create Catalog
        ////    var catalogLayout = new PdfCatalogLayout(_workheets[0], _pageSettings, _dictionaries);
        ////    var catalog = AddCatalog(2);
        ////    //Create Pages
        ////    var pagesLayout = catalogLayout.ChildObjects[0];
        ////    var pages = AddPages();
        ////    //Create Fonts
        ////    AddFontData();
        ////    //Create Patterns
        ////    AddPatternData();
        ////    //Create Shadings
        ////    AddShadingsData();
        ////    //Create Page and Content
        ////    for (int i = 0; i < pagesLayout.ChildObjects.Count; i++)
        ////    {
        ////        var pageLayout = pagesLayout.ChildObjects[i];
        ////        var page = AddPage(2, new List<int>(), _pageSettings);
        ////        AddContent(pageLayout, page);
        ////        pages.pageObjectNumbers.Add(page.objectNumber);
        ////    }
        ////    var info = AddInfoObject(_workheets[0].Workbook._package.File.Name);

        ////    WriteDocumentToStream(stream, catalog, info);
        ////}

        ////Write the document and cross-ref/trailer to the supplied stream.
        ////The stream is not closed; the caller owns it.
        //private void WriteDocumentToStream(Stream stream, PdfObjects.PdfCatalog catalog, PdfInfoObject info)
        //{
        //    _debugString = "";
        //    PdfCrossRefTable crossRefTable = new PdfCrossRefTable();

        //    //Use a BinaryWriter without disposing it, so the underlying stream stays open for the caller.
        //    //BinaryWriter does not own the stream when we don't dispose it; we just flush at the end.
        //    var bw = new BinaryWriter(stream, Encoding.ASCII);
        //    try
        //    {
        //        //Write header
        //        bw.Write(Encoding.ASCII.GetBytes(Header));
        //        _debugString += Header;
        //        //Write body
        //        foreach (var pdfobj in _document)
        //        {
        //            crossRefTable.AddPosition(stream.Position);
        //            pdfobj.ToPdfBytes(bw);
        //            _debugString += pdfobj.ToPdfString();
        //        }
        //        //Write CrossReference
        //        crossRefTable.Write(bw, stream.Position, _document.Count);
        //        _debugString += crossRefTable.WriteString(_document.Count);
        //        // Write trailer
        //        PdfTrailer.Write(bw, _document.Count, catalog.objectNumber, info.objectNumber, crossRefTable.StartPosition);
        //        _debugString += PdfTrailer.WriteString(_document.Count, catalog.objectNumber, info.objectNumber, crossRefTable.StartPosition);
        //    }
        //    finally
        //    {
        //        bw.Flush();
        //    }
        //}

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