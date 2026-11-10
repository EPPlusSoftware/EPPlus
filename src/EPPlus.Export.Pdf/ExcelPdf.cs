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

        //Add Image Data
        private void AddImageData()
        {
            foreach (var image in _dictionaries.Images)
            {
                var img = image.Value.GetImageObject(_document.Count + 1, _documentSettings.BlackAndWhite);
                if (img.HasSoftMask)
                {
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
        private void AddContent(PdfPageLayout pageLayout, PdfPage page, List<(PdfPageLayout Layout, int PageObj)> destIndex)
        {
            var pageSettings = pageLayout.Settings;
            double contentScale = pageSettings.ContentScale;
            double scaleAnchorX = pageSettings.ContentBounds.Left;
            double scaleAnchorY = pageSettings.ContentBounds.Top;
            var cells = pageLayout.ChildObjects.Where(t =>
                                                     (t is PdfCellLayout || t is PdfCellContentLayout || t is PdfCellBorderLayout) &&
                                                    !(t is PdfCellLayout cc && (cc.IsHeading || cc.IsPrintTitle)) &&
                                                    !(t is PdfCellContentLayout ccl && (ccl.IsHeaderFooter || ccl.IsHeading || ccl.IsPrintTitle || ccl.IsAddedContent)) &&
                                                    !(t is PdfCellBorderLayout cbl && cbl.IsPrintTitle)).ToList();
            var AdditionalContent = pageLayout.ChildObjects.OfType<PdfCellContentLayout>().Where(t => t.IsAddedContent && !(pageSettings.AdditionalContent != "null"));
            var headingLayouts = pageLayout.ChildObjects.Where(t => (t is PdfCellLayout cl && cl.IsHeading) || (t is PdfCellContentLayout ccl && ccl.IsHeading));
            var printTitleLayouts = pageLayout.ChildObjects.Where(t => (t is PdfCellLayout pl && pl.IsPrintTitle) || (t is PdfCellContentLayout pcl && pcl.IsPrintTitle) || (t is PdfCellBorderLayout pbl && pbl.IsPrintTitle));
            var contentStream = new PdfContentStream(_document.Count + 1);
            //Start page content clipping rectangle.
            contentStream.AddCommand("q");
            contentStream.AddContentScale(contentScale, scaleAnchorX, scaleAnchorY);
            contentStream.AddMarginClipping((PdfPageLayout)pageLayout, pageSettings);
            if (pageSettings.ShowGridLines)
            {
                contentStream.AddInnerGridLines(pageLayout);
            }
            foreach (PdfCellLayout fill in cells.OfType<PdfCellLayout>())
            {
                contentStream.AddCellLayout(fill, GetPatternLabel(fill));
            }
            foreach (PdfCellContentLayout content in cells.OfType<PdfCellContentLayout>())
            {
                contentStream.AddCellContentLayout(content, _dictionaries, pageSettings);
            }
            foreach (PdfCellBorderLayout border in cells.OfType<PdfCellBorderLayout>())
            {
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
            //Add headings
            contentStream.AddCommand("q");
            contentStream.AddContentScale(contentScale, scaleAnchorX, scaleAnchorY);
            foreach (var heading in headingLayouts)
            {
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
            contentStream.AddCommand("Q");
            //Add outer gridlines
            if (pageSettings.ShowGridLines || pageSettings.ShowHeadings)
            {
                contentStream.AddCommand("q");
                contentStream.AddContentScale(contentScale, scaleAnchorX, scaleAnchorY);
                contentStream.AddOuterGridBorder(pageLayout);
                contentStream.AddPrintTitleGridLines(pageLayout);
                contentStream.AddCommand("Q");
            }
            //Add header and footer.
            foreach (var part in pageLayout.HeaderFooterParts)
            {
                bool scaleHf = pageLayout.ScaleHeaderFooterWithDocument;
                if (scaleHf)
                {
                    contentStream.AddCommand("q");
                    contentStream.AddContentScale(contentScale, part.ScalePivot.X, part.ScalePivot.Y);
                }

                foreach (var item in part.Items)
                {
                    switch (item)
                    {
                        case PdfCellContentLayout text:
                            contentStream.AddCellContentLayout(text, _dictionaries, pageSettings);
                            break;
                        case PdfImageLayout image:
                            var imageResource = _dictionaries.AddImage(image.ImageBytes);
                            contentStream.AddImage(imageResource.Label, image.LocalPosition.X, image.LocalPosition.Y, image.Size.X, image.Size.Y);
                            if (PdfImageXObject.ProducesSoftMask(image.ImageBytes)) page.HasTransparency = true;
                            break;
                    }
                }

                if (scaleHf)
                    contentStream.AddCommand("Q");
            }
            //Add print titles
            contentStream.AddCommand("q");
            contentStream.AddContentScale(contentScale, scaleAnchorX, scaleAnchorY);
            foreach (var titleCell in printTitleLayouts)
            {
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
            contentStream.AddCommand("Q");
            foreach (var content in AdditionalContent)
            {
                contentStream.AddCellContentLayout(content, _dictionaries, pageSettings);
            }
            _document.Add(contentStream);
            page.contentObjectNumbers.Add(contentStream.objectNumber);
            foreach (var link in pageLayout.Links)
            {
                double lowerLeftX = contentScale * link.X + (1 - contentScale) * scaleAnchorX;
                double lowerLeftY = contentScale * link.Y + (1 - contentScale) * scaleAnchorY;
                double upperRightX = contentScale * (link.X + link.Width) + (1 - contentScale) * scaleAnchorX;
                double upperRightY = contentScale * (link.Y + link.Height) + (1 - contentScale) * scaleAnchorY;

                PdfLinkAnnotationObject annot;
                if (!string.IsNullOrEmpty(link.Uri))
                {
                    annot = new PdfLinkAnnotationObject(_document.Count + 1, lowerLeftX, lowerLeftY, upperRightX, upperRightY, link.Uri, link.ToolTip);
                }
                else if (link.TargetSheet != null && TryResolveDest(destIndex, link.TargetSheet, link.TargetRow, link.TargetColumn, out int destPageObj, out double destX, out double destY))
                {
                    annot = PdfLinkAnnotationObject.GoTo(_document.Count + 1,lowerLeftX, lowerLeftY, upperRightX, upperRightY, destPageObj, destX, destY, link.ToolTip);
                }
                else
                {
                    continue;
                }

                _document.Add(annot);
                page.annotationObjectNumbers.Add(annot.objectNumber);
            }
        }

        private void AddNamedDestinations(PdfCatalog catalog, List<(PdfPageLayout Layout, int PageObj)> destIndex, PdfCatalogLayout catLayout)
        {
            if (catLayout == null || catLayout.NamedDestinations.Count == 0) return;

            var namesObj = new PdfNamedDestinations(_document.Count + 1);
            var seenDest = new HashSet<string>(StringComparer.Ordinal);
            foreach (var nd in catLayout.NamedDestinations)
            {
                if (seenDest.Contains(nd.Name)) continue;
                if (TryResolveDest(destIndex, nd.Sheet, nd.Row, nd.Col, out int po, out double dx, out double dy))
                {
                    namesObj.Dests.Add((nd.Name, po, dx, dy));
                    seenDest.Add(nd.Name);
                }
            }
            if (namesObj.Dests.Count > 0)
            {
                _document.Add(namesObj);
                catalog.NamesObjectNumber = namesObj.objectNumber;
            }
        }

        private void AddOutline(PdfCatalog catalog, List<(PdfPageLayout Layout, int PageObj)> destIndex, PdfCatalogLayout catLayout)
        {
            var pageOrder = new Dictionary<int, int>();
            for (int i = 0; i < destIndex.Count; i++)
                pageOrder[destIndex[i].PageObj] = i;
            var sheets = new List<(string Name, int PageObj, double TopY)>();
            var seenSheets = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var commentsBySheet = new Dictionary<string, (int PageObj, double TopY)>(StringComparer.OrdinalIgnoreCase);
            foreach (var (L, obj) in destIndex)
            {
                if (!L.isCommentsPage && L.SheetName != null)
                {
                    if (seenSheets.Add(L.SheetName))
                        sheets.Add((L.SheetName, obj, L.Settings.PageSize.HeightPu));
                }
                else if (L.isCommentsPage && L.CommentsSourceSheet != null)
                {
                    if (!commentsBySheet.ContainsKey(L.CommentsSourceSheet))
                        commentsBySheet[L.CommentsSourceSheet] = (obj, L.Settings.PageSize.HeightPu);
                }
            }
            if (sheets.Count == 0) return;
            var outlines = new PdfOutlines(_document.Count + 1);
            _document.Add(outlines);
            var sheetItems = new List<PdfOutlineItem>();
            int totalVisible = 0;
            foreach (var sh in sheets)
            {
                var sheetItem = new PdfOutlineItem(_document.Count + 1)
                {
                    Title = sh.Name,
                    Parent = outlines.objectNumber,
                    DestPageObj = sh.PageObj,
                    DestX = 0d,
                    DestY = sh.TopY
                };
                _document.Add(sheetItem);
                sheetItems.Add(sheetItem);
                totalVisible++;
                var childItems = new List<PdfOutlineItem>();
                foreach (var k in BuildOutlineChildren(sh.Name, _documentSettings.BookmarkNamePrefix, catLayout, commentsBySheet, destIndex, pageOrder))
                {
                    var childItem = new PdfOutlineItem(_document.Count + 1)
                    {
                        Title = k.Title,
                        Parent = sheetItem.objectNumber,
                        DestPageObj = k.PageObj,
                        DestX = k.X,
                        DestY = k.Y
                    };
                    _document.Add(childItem);
                    childItems.Add(childItem);
                    totalVisible++;
                }
                if (childItems.Count > 0)
                {
                    for (int i = 0; i < childItems.Count; i++)
                    {
                        childItems[i].Prev = i > 0 ? childItems[i - 1].objectNumber : 0;
                        childItems[i].Next = i < childItems.Count - 1 ? childItems[i + 1].objectNumber : 0;
                    }
                    sheetItem.First = childItems[0].objectNumber;
                    sheetItem.Last = childItems[childItems.Count - 1].objectNumber;
                    sheetItem.Count = childItems.Count;
                }
            }
            for (int i = 0; i < sheetItems.Count; i++)
            {
                sheetItems[i].Prev = i > 0 ? sheetItems[i - 1].objectNumber : 0;
                sheetItems[i].Next = i < sheetItems.Count - 1 ? sheetItems[i + 1].objectNumber : 0;
            }
            outlines.First = sheetItems[0].objectNumber;
            outlines.Last = sheetItems[sheetItems.Count - 1].objectNumber;
            outlines.Count = totalVisible;
            catalog.OutlinesObjectNumber = outlines.objectNumber;
        }

        private static List<(string Title, int PageObj, double X, double Y)> BuildOutlineChildren(string sheetName, string prefix, PdfCatalogLayout catLayout, Dictionary<string, (int PageObj, double TopY)> commentsBySheet, List<(PdfPageLayout Layout, int PageObj)> destIndex, Dictionary<int, int> pageOrder)
        {
            bool hasPrefix = !string.IsNullOrEmpty(prefix);
            var kids = new List<(string Title, int PageObj, double X, double Y)>();
            if (catLayout != null)
            {
                foreach (var nd in catLayout.NamedDestinations)
                {
                    if (nd.Name == null) continue;
                    if (!string.Equals(nd.Sheet, sheetName, StringComparison.OrdinalIgnoreCase)) continue;
                    if (hasPrefix && !nd.Name.StartsWith(prefix, StringComparison.OrdinalIgnoreCase)) continue;
                    string namePart = hasPrefix ? nd.Name.Substring(prefix.Length) : nd.Name;
                    string pretty = namePart.Replace('_', ' ').Trim();
                    string title = !string.IsNullOrWhiteSpace(nd.CellText) ? nd.CellText : pretty;
                    if (string.IsNullOrWhiteSpace(title)) continue;
                    if (!TryResolveDest(destIndex, nd.Sheet, nd.Row, nd.Col, out int po, out double dx, out double dy)) continue;
                    kids.Add((title, po, dx, dy));
                }
            }
            if (commentsBySheet.TryGetValue(sheetName, out var cm))
                kids.Add(("Comments & Notes", cm.PageObj, 0d, cm.TopY));

            kids.Sort((a, b) =>
            {
                int pa = pageOrder.TryGetValue(a.PageObj, out var va) ? va : int.MaxValue;
                int pb = pageOrder.TryGetValue(b.PageObj, out var vb) ? vb : int.MaxValue;
                if (pa != pb) return pa.CompareTo(pb);
                if (a.Y != b.Y) return b.Y.CompareTo(a.Y);
                return a.X.CompareTo(b.X);
            });
            return kids;
        }

        private static bool TryResolveDest(List<(PdfPageLayout Layout, int PageObj)> index, string sheet, int row, int col, out int pageObj, out double x, out double y)
        {
            pageObj = 0; x = 0; y = 0;
            foreach (var (L, obj) in index)
            {
                if (L.ColLeftX == null || L.SheetName == null) continue;
                if (!string.Equals(L.SheetName, sheet, StringComparison.OrdinalIgnoreCase)) continue;
                if (row < L.FromRow || row > L.ToRow || col < L.FromColumn || col > L.ToColumn) continue;

                double scale = L.Settings.ContentScale;
                x = scale * L.ColLeftX[col - L.FromColumn] + (1 - scale) * L.Settings.ContentBounds.Left;
                y = scale * L.RowTopY[row - L.FromRow] + (1 - scale) * L.Settings.ContentBounds.Top;
                pageObj = obj;
                return true;
            }
            return false;
        }

        //Add Info
        private PdfInfoObject AddInfoObject()
        {
            var info = new PdfInfoObject(_document.Count + 1, _documentSettings);
            _document.Add(info);
            return info;
        }

        //Creates the Pdf document
        internal void CreatePdf(PdfDocumentSettings documentSettings, PdfDictionaries dictionaries, Transform layout, string fileName)
        {
            //Write the PDF to the file.
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

        //Creates the Pdf document
        internal void CreatePdf(PdfDocumentSettings documentSettings, PdfDictionaries dictionaries, Transform layout, Stream stream)
        {
            if (stream == null) throw new ArgumentNullException(nameof(stream));
            if (!stream.CanWrite) throw new ArgumentException("The stream must be writable.", nameof(stream));
            //The cross-reference table stores byte offsets that the PDF reader uses to
            //seek to each object, so the target stream has to support querying its position.
            if (!stream.CanSeek) throw new ArgumentException("The stream must be seekable, because the PDF cross-reference table requires byte offsets.", nameof(stream));
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
            var pageObjs = new PdfPage[layout.ChildObjects.Count];
            var destIndex = new List<(PdfPageLayout Layout, int PageObj)>();
            for (int i = 0; i < layout.ChildObjects.Count; i++)
            {
                var pageLayout = (PdfPageLayout)layout.ChildObjects[i];
                pageObjs[i] = AddPage(2, new List<int>(), pageLayout.Settings);
                destIndex.Add((pageLayout, pageObjs[i].objectNumber));
                pages.pageObjectNumbers.Add(pageObjs[i].objectNumber);
            }
            for (int i = 0; i < layout.ChildObjects.Count; i++)
            {
                AddContent((PdfPageLayout)layout.ChildObjects[i], pageObjs[i], destIndex);
            }
            var catLayout = layout as PdfCatalogLayout;
            //Create Named destinations
            AddNamedDestinations(catalog, destIndex, catLayout);
            //Create Bookmark outlines
            AddOutline(catalog, destIndex, catLayout);
            //Create Image data
            AddImageData();
            var info = AddInfoObject();
            _debugString = "";
            //Write to pdf
            PdfCrossRefTable crossRefTable = new PdfCrossRefTable();
            //Cross-reference offsets are relative to the start of the PDF.
            //A freshly created FileStream starts at position 0, but a caller-supplied stream may already hold data, so capture the starting position and make every offset relative to it.
            long start = stream.Position;
            //Start writing pdf binary. leaveOpen: true so a caller-supplied stream is not closed.
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