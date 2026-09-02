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
using EPPlus.Export.Pdf.DocumentObjects.Functions;
using EPPlus.Export.Pdf.Enums;
using EPPlus.Export.Pdf.Layout;
using EPPlus.Export.Pdf.Resources;
using EPPlus.Export.Pdf.Settings;
using EPPlus.Graphics;
using System;
using System.Collections.Generic;
using System.Diagnostics;
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
        private PdfDocumentSettings _documentSettings; 
        private PdfDictionaries _dictionaries;
        internal List<PdfObject> _document = new List<PdfObject>();
        private string _debugString;

        internal static string Header
        {
            get
            {
                return "%PDF-1.7\n";
            }
        }

        internal void SetPageSettingsForTest(PdfPageSettings pageSettings)
        {
            _pageSettings = pageSettings;
        }

        internal void SetDictionariesForTest(PdfDictionaries dictionaries)
        {
            _dictionaries = dictionaries;
        }

        internal void SetDocumentSettingsForTest(PdfDocumentSettings documentSettings)
        {
            _documentSettings = documentSettings;
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
        internal void AddFontData()
        {
            foreach (var f in _dictionaries.Fonts)
                Debug.WriteLine($"Fonts: {f.Key} → label={f.Value.Label} nr={f.Value.labelNumber}");
            if (_documentSettings.EmbeddFonts)
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
                    font.Value.fontObjectNumber = font.Value.type0FontObjectNumber;
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
                var gradient = shading.Value.CellFillData.GradientFillData;
                if (gradient != null && gradient.GradientType == ExcelFillGradientType.Path)
                {
                    // Box gradient: ShadingType 1 + Type 4 PostScript function. A Type 4 function is
                    // a stream object, so it must be its own indirect object referenced by the shading.
                    var boxFunction = new PdfPostScriptCalculatorFunction(_document.Count + 1, gradient);
                    _document.Add(boxFunction);
                    _document.Add(shading.Value.GetShadingObject(_document.Count + 1, boxFunction.objectNumber));
                }
                else
                {
                    _document.Add(shading.Value.GetShadingObject(_document.Count + 1));
                }
                _document.Add(shading.Value.GetShadingPatternObject(_document.Count + 1, _document.Count));
                int label = _dictionaries.Patterns.Last().Value.labelNumber + 1;
                var pr = new PdfPatternResource(label, shading.Value.CellFillData);
                pr.objectNumber = _document.Count;
                _dictionaries.Patterns.Add(shading.Value.CellFillData.id, pr);
            }
        }

        //Add Images
        private void AddImageData()
        {
            foreach (var image in _dictionaries.Images)
            {
                var img = image.Value.GetImageObject(_document.Count + 1);
                if (img.HasSoftMask)
                {
                    // Alpha PNG: the alpha channel is a separate grayscale /SMask object. Add it
                    // first, then point the image at it and shift the image (and the page /XObject
                    // reference in image.Value) to the next slot so all three numbers agree.
                    var mask = PdfImageXObject.CreateSoftMask(_document.Count + 1, img.SoftMaskData, img.Width, img.Height);
                    _document.Add(mask);
                    img.SoftMaskObjectNumber = mask.objectNumber;
                    img.objectNumber = _document.Count + 1;
                    image.Value.objectNumber = img.objectNumber;
                }
                _document.Add(img);
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
        private void AddContent(PdfPageLayout pageLayout, PdfPage page)
        {
            var pageSettings = pageLayout.Settings;

            var cells = pageLayout.ChildObjects.Where(t =>
                                                     (t is PdfCellLayout || t is PdfCellContentLayout || t is PdfCellBorderLayout) &&
                                                    !(t is PdfCellLayout cc && (cc.IsHeading || cc.IsPrintTitle)) &&
                                                    !(t is PdfCellContentLayout ccl && (ccl.IsHeaderFooter || ccl.IsHeading || ccl.IsPrintTitle)) &&
                                                    !(t is PdfCellBorderLayout cbl && cbl.IsPrintTitle)).ToList();

            var headerFooterLayouts = pageLayout.ChildObjects.OfType<PdfCellContentLayout>().Where(t => t.IsHeaderFooter);
            var headingLayouts = pageLayout.ChildObjects.Where(t => (t is PdfCellLayout cl && cl.IsHeading) || (t is PdfCellContentLayout ccl && ccl.IsHeading));
            var printTitleLayouts = pageLayout.ChildObjects.Where(t => (t is PdfCellLayout pl && pl.IsPrintTitle) || (t is PdfCellContentLayout pcl && pcl.IsPrintTitle) || (t is PdfCellBorderLayout pbl && pbl.IsPrintTitle));
            var contentStream = new PdfContentStream(_document.Count + 1);
            contentStream.AddCommand($"% {pageLayout.Name} start");
            //Add clipping rectangle around page content.
            contentStream.AddCommand("q");
            contentStream.AddMarginClipping((PdfPageLayout)pageLayout, pageSettings);
            if (pageSettings.ShowGridLines)
            {
                contentStream.AddInnerGridLines(pageLayout);
            }
            foreach (PdfCellLayout fill in cells.OfType<PdfCellLayout>())
            {
                contentStream.AddCommand($"% CELL FILL : {fill.Name}");
                contentStream.AddCellLayout(fill, GetPatternLabel(fill));
            }
            foreach (PdfCellContentLayout content in cells.OfType<PdfCellContentLayout>())
            {
                contentStream.AddCommand($"% CELL TEXT : {content.Name}");
                contentStream.AddCellContentLayout(content, _dictionaries, pageSettings);
            }
            foreach (PdfCellBorderLayout border in cells.OfType<PdfCellBorderLayout>())
            {
                contentStream.AddCommand($"% CELL BORDER : {border.Name}");
                contentStream.AddBorderLayout(border);
            }
            foreach (PdfImageLayout image in pageLayout.ChildObjects.OfType<PdfImageLayout>())
            {
                if (image.IsHeaderFooter) continue;
                var imageResource = _dictionaries.AddImage(image.ImageBytes);
                contentStream.AddImage(imageResource.Label, image.LocalPosition.X, image.LocalPosition.Y, image.Size.X, image.Size.Y);
                if (PdfImageXObject.ProducesSoftMask(image.ImageBytes)) page.HasTransparency = true;
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
                        contentStream.AddCellContentLayout(contentLayout, _dictionaries, pageSettings); break;
                    case PdfCellBorderLayout borderLayout:
                        contentStream.AddBorderLayout(borderLayout); break;
                }
            }
            if (pageSettings.ShowGridLines || pageSettings.ShowHeadings)
            {
                contentStream.AddOuterGridBorder(pageLayout);
                contentStream.AddPrintTitleGridLines(pageLayout);
            }
            //Add header and footer.
            foreach (var hf in headerFooterLayouts)
            {
                contentStream.AddCellContentLayout(hf, _dictionaries, pageSettings);
            }
            foreach (PdfImageLayout image in pageLayout.ChildObjects.OfType<PdfImageLayout>())
            {
                if (!image.IsHeaderFooter) continue;
                var imageResource = _dictionaries.AddImage(image.ImageBytes);
                contentStream.AddImage(imageResource.Label, image.LocalPosition.X, image.LocalPosition.Y, image.Size.X, image.Size.Y);
                if (PdfImageXObject.ProducesSoftMask(image.ImageBytes)) page.HasTransparency = true;
            }
            foreach (var titleCell in printTitleLayouts)
            {
                contentStream.AddCommand($"% PRINT TITLE : {titleCell.Name}");
                switch (titleCell)
                {
                    case PdfCellLayout layout:
                        contentStream.AddCellLayout(layout, GetPatternLabel(layout)); break;
                    case PdfCellContentLayout contentLayout:
                        contentStream.AddCellContentLayout(contentLayout, _dictionaries, pageSettings); break;
                    case PdfCellBorderLayout borderLayout:
                        contentStream.AddBorderLayout(borderLayout); break;
                }
            }
            _document.Add(contentStream);
            page.contentObjectNumbers.Add(contentStream.objectNumber);
            contentStream.AddCommand($"% {pageLayout.Name} end");
        }

        //Add Info
        private PdfInfoObject AddInfoObject(string workBookName = "")
        {
            var info = new PdfInfoObject(_document.Count + 1, workBookName);
            _document.Add(info);
            return info;
        }

        internal void CreatePdf(PdfDocumentSettings documentSettings, PdfDictionaries dictionaries, Transform layout, string fileName)
        {
            //Write the PDF to the file. The Stream overload does the actual work and
            //populates _debugString.
            using (var fs = new FileStream(fileName, FileMode.Create, FileAccess.Write))
            {
                CreatePdf(documentSettings, dictionaries, layout, fs);
            }
            //Write pdf as txt for debug.
            if (_documentSettings.Debug && _documentSettings.PrintAsText)
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

        internal void CreatePdf(PdfDocumentSettings documentSettings, PdfDictionaries dictionaries, Transform layout, Stream stream)
        {
            if (stream == null) throw new ArgumentNullException(nameof(stream));
            if (!stream.CanWrite) throw new ArgumentException("The stream must be writable.", nameof(stream));
            //The cross-reference table stores byte offsets that the PDF reader uses to
            //seek to each object, so the target stream has to support querying its position.
            if (!stream.CanSeek) throw new ArgumentException("The stream must be seekable, because the PDF cross-reference table requires byte offsets.", nameof(stream));

            //_pageSettings = pageSettings;
            _documentSettings = documentSettings;
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
                var pageLayout = (PdfPageLayout)layout.ChildObjects[i];  
                var page = AddPage(2, new List<int>(), pageLayout.Settings);
                AddContent(pageLayout, page);
                pages.pageObjectNumbers.Add(page.objectNumber);
            }
            AddImageData();
            var info = AddInfoObject();
            _debugString = "";
            //write to pdf
            PdfCrossRefTable crossRefTable = new PdfCrossRefTable();
            //Cross-reference offsets are relative to the start of the PDF. A freshly created
            //FileStream starts at position 0, but a caller-supplied stream may already hold
            //data, so capture the starting position and make every offset relative to it.
            long start = stream.Position;
            //start writing pdf binary. leaveOpen: true so a caller-supplied stream is not closed.
            using (var bw = new BinaryWriter(stream, Encoding.ASCII, true))
            {
                //Write header
                bw.Write(Encoding.ASCII.GetBytes(Header));
                _debugString += Header;
                //Write body
                foreach (var pdfobj in _document)
                {
                    crossRefTable.AddPosition(stream.Position - start);
                    pdfobj.ToPdfBytes(bw);
                    _debugString += pdfobj.ToPdfString();
                }
                //Write CrossReference
                crossRefTable.Write(bw, stream.Position - start, _document.Count);
                _debugString += crossRefTable.WriteString(_document.Count);
                // Write trailer
                PdfTrailer.Write(bw, _document.Count, catalog.objectNumber, info.objectNumber, crossRefTable.StartPosition);
                _debugString += PdfTrailer.WriteString(_document.Count, catalog.objectNumber, info.objectNumber, crossRefTable.StartPosition);
                bw.Flush();
            }
        }
    }
}