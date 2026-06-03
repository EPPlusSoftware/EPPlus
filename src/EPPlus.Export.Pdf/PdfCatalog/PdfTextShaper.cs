using EPPlus.Export.Pdf.PdfLayout;
using EPPlus.Export.Pdf.PdfResources;
using EPPlus.Export.Pdf.PdfSettings;
using EPPlus.Fonts.OpenType;
using EPPlus.Fonts.OpenType.Integration;
using EPPlus.Fonts.OpenType.TextShaping;
using OfficeOpenXml.Interfaces.Drawing.Text;
using OfficeOpenXml.Interfaces.Fonts;
using System;
using System.Collections.Generic;
using System.Linq;

namespace EPPlus.Export.Pdf.PdfCatalog
{
    internal static class PdfTextShaper
    {
        private static Dictionary<IFontProvider, TextShaper> shaperCache = new Dictionary<IFontProvider, TextShaper>();
        private static Dictionary<IFontProvider, TextLayoutEngine> layoutEngineCache = new Dictionary<IFontProvider, TextLayoutEngine>();

        // Pass 1: collect text per font so FontSubsetManager can build subsets once
        public static void CollectText(PdfDictionaries dictionaries, PdfCell cell)
        {
            if (cell == null || cell.TextFragments == null) return;

            for (int i = 0; i < cell.TextFragments.Count; i++)
            {
                var tf = cell.TextFragments[i];
                if (!dictionaries.Fonts.ContainsKey(tf.FullFontName)) continue;
                dictionaries.Fonts[tf.FullFontName].fontSubsetManager.AddText(tf.Text);
            }
        }

        // Pass 2: shape text using already-built providers from PdfDictionaries.ShapedProviders
        public static void ShapeText(PdfPageSettings pageSettings, PdfDictionaries dictionaries, PdfCell cell)
        {
            var totalTextLength = 0d;
            var maxLineHeight = 0d;
            if (cell == null || cell.TextFragments == null) return;

            cell.ShapedTexts = new List<PdfShapedText>();

            for (int i = 0; i < cell.TextFragments.Count; i++)
            {
                var tf = cell.TextFragments[i];
                cell.ShapedTexts.Add(new PdfShapedText());
                var st = cell.ShapedTexts[i];

                if (!dictionaries.ShapedProviders.TryGetValue(tf.FullFontName, out var provider))
                {
                    continue;
                }
                st.FontProvider = provider;

                if (!shaperCache.TryGetValue(st.FontProvider, out var shaper))
                {
                    shaper = new TextShaper(st.FontProvider);
                    shaperCache[st.FontProvider] = shaper;
                }

                if (!layoutEngineCache.TryGetValue(st.FontProvider, out var layoutEngine))
                {
                    layoutEngine = new TextLayoutEngine(shaper);
                    layoutEngineCache[st.FontProvider] = layoutEngine;
                }

                var options = ShapingOptions.Default;
                options.ApplyPositioning = true;
                options.ApplySubstitutions = true;

                var shaped = shaper.Shape(tf.Text, options);
                var usedFonts = shaper.GetUsedFonts().ToList();
                //Console.WriteLine($"[ShapeText] tf.Text='{tf.Text}' FullFontName='{tf.FullFontName}' usedFonts={string.Join(",", usedFonts.Select(f => f.FullName).ToArray())}");
                var fontIdMap = new Dictionary<byte, string>();

                for (byte fontId = 0; fontId < usedFonts.Count; fontId++)
                {
                    var font = usedFonts[fontId];
                    if (!dictionaries.Fonts.ContainsKey(font.FullName))
                    {
                        int label = dictionaries.Fonts.Count > 0
                            ? dictionaries.Fonts.Last().Value.labelNumber + 1
                            : 1;
                        var fontResource = new PdfFontResource(font.FullName, font.NameTable.GetSubfamilyEnum(), label, pageSettings);
                        fontResource.fontData = font;
                        dictionaries.Fonts.Add(font.FullName, fontResource);
                    }
                    fontIdMap[fontId] = dictionaries.Fonts[font.FullName].Label;
                }

                cell.TextLayoutEngine = layoutEngine;
                st.ShapedText = shaped;

                totalTextLength += st.ShapedText.GetWidthInPoints((float)tf.Font.Size);
                maxLineHeight = Math.Max(st.ShapedText.GetLineHeightInPoints((float)tf.Font.Size), maxLineHeight);

                st.FontIdMap = fontIdMap;
                st.UsedFonts = usedFonts;
                cell.TextFragments[i] = tf;
                cell.ShapedTexts[i] = st;
            }

            if (cell.TextLayoutEngine != null)
            {
                double wrapWidth = (cell.Merged && cell.Main == null) ? cell.Width : cell.ColumnWidth;
                cell.TextLines = cell.ContentAligmnet.WrapText
                    ? cell.TextLayoutEngine.WrapRichTextLineCollection(cell.TextFragments, wrapWidth)
                    : cell.TextLayoutEngine.WrapRichTextLineCollection(cell.TextFragments, double.MaxValue);
            }

            cell.TotalTextLength = totalTextLength;
        }

        // Kept for backwards compatibility - not used by ShapeTextInPdfWorksheet anymore
        public static void LayoutAndShapeText(PdfPageSettings pageSettings, PdfDictionaries dictionaries, PdfCell cell)
        {
            var totalTextLength = 0d;
            var maxLineHeight = 0d;
            if (cell.TextFragments == null) return;
            cell.ShapedTexts = new List<PdfShapedText>();
            for (int i = 0; i < cell.TextFragments.Count; i++)
            {
                var tf = cell.TextFragments[i];
                cell.ShapedTexts.Add(new PdfShapedText());
                var st = cell.ShapedTexts[i];
                st.FontProvider = dictionaries.Fonts[tf.FullFontName].fontSubsetManager.CreateSubsettedProvider();

                if (!shaperCache.TryGetValue(st.FontProvider, out var shaper))
                {
                    shaper = new TextShaper(st.FontProvider);
                    shaperCache[st.FontProvider] = shaper;
                }

                if (!layoutEngineCache.TryGetValue(st.FontProvider, out var layoutEngine))
                {
                    layoutEngine = new TextLayoutEngine(shaper);
                    layoutEngineCache[st.FontProvider] = layoutEngine;
                }

                var options = ShapingOptions.Default;
                options.ApplyPositioning = true;
                options.ApplySubstitutions = true;

                var shaped = shaper.Shape(tf.Text, options);
                var usedFonts = shaper.GetUsedFonts().ToList();
                var fontIdMap = new Dictionary<byte, string>();

                var allProviderFonts = st.FontProvider.GetAllFonts().ToList();

                for (byte fontId = 0; fontId < usedFonts.Count; fontId++)
                {
                    var font = usedFonts[fontId];
                    if (!dictionaries.Fonts.ContainsKey(font.FullName))
                    {
                        int label = 1;
                        if (dictionaries.Fonts.Count > 0)
                        {
                            label = dictionaries.Fonts.Last().Value.labelNumber + 1;
                        }
                        var fontResource = new PdfFontResource(font.FullName, font.NameTable.GetSubfamilyEnum(), label, pageSettings);
                        fontResource.fontData = font;
                        dictionaries.Fonts.Add(font.FullName, fontResource);
                    }
                    fontIdMap[fontId] = dictionaries.Fonts[font.FullName].Label;
                }
                cell.TextLayoutEngine = layoutEngine;
                st.ShapedText = shaped;
                var textWidth = st.ShapedText.GetWidthInPoints((float)tf.Font.Size);
                var textHeight = st.ShapedText.GetLineHeightInPoints((float)tf.Font.Size);
                totalTextLength += textWidth;
                maxLineHeight = Math.Max(textHeight, maxLineHeight);
                st.FontIdMap = fontIdMap;
                st.UsedFonts = usedFonts;
                cell.TextFragments[i] = tf;
                cell.ShapedTexts[i] = st;
            }
            if (cell.TextLayoutEngine != null)
            {
                if (cell.ContentAligmnet.WrapText)
                {
                    double wrapWidth = (cell.Merged && cell.Main == null) ? cell.Width : cell.ColumnWidth;
                    cell.TextLines = cell.TextLayoutEngine.WrapRichTextLineCollection(cell.TextFragments, wrapWidth);
                }
                else
                {
                    cell.TextLines = cell.TextLayoutEngine.WrapRichTextLineCollection(cell.TextFragments, double.MaxValue);
                }
            }
            cell.TotalTextLength = totalTextLength;
        }
    }
}