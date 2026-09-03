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
using EPPlus.Fonts.OpenType;
using EPPlus.Fonts.OpenType.Integration;
using EPPlus.Fonts.OpenType.TextShaping;
using OfficeOpenXml.Export.PdfExport.Data;
using OfficeOpenXml.Interfaces.Fonts;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;

namespace OfficeOpenXml.Export.PdfExport.TextShaping
{
    internal static class PdfTextShaper
    {
        private static Dictionary<IFontProvider, TextShaper> shaperCache = new Dictionary<IFontProvider, TextShaper>();
        private static Dictionary<IFontProvider, TextLayoutEngine> layoutEngineCache = new Dictionary<IFontProvider, TextLayoutEngine>();

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
                var key = dictionaries.ResolveFontKey(pageSettings, tf.Font.Family, tf.Font.SubFamily);
                IFontProvider provider;
                if (!dictionaries.ShapedProviders.TryGetValue(key, out provider))
                {
                    // No subset provider was built for this font — this is the measurement path
                    // (GetCellCollectionFromRange), which does not run BuildSubsets. Shape against the
                    // whole font instead: advance widths are identical to the subset, so measured width
                    // is exact, and no subsetting or embedding decision is triggered.
                    var font = pageSettings.FontEngine.LoadFont(tf.Font.Family, tf.Font.SubFamily);
                    provider = new DefaultFontProvider(pageSettings.FontEngine, font);
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
                var fontIdMap = new Dictionary<byte, string>();
                for (byte fontId = 0; fontId < usedFonts.Count; fontId++)
                {
                    var font = usedFonts[fontId];
                    var loadedKey = new FontKey(font.GetEnglishFontFamilyName(), font.NameTable.GetSubfamilyEnum());
                    if (!dictionaries.Fonts.ContainsKey(loadedKey))
                    {
                        int label = dictionaries.Fonts.Count > 0
                            ? dictionaries.Fonts.Last().Value.labelNumber + 1
                            : 1;
                        var fontResource = new PdfFontResource(font.GetEnglishFontFamilyName(), font.NameTable.GetSubfamilyEnum(), label, pageSettings);
                        fontResource.fontData = font;
                        dictionaries.Fonts.Add(loadedKey, fontResource);
                    }
                    fontIdMap[fontId] = dictionaries.Fonts[loadedKey].Label;
                }
                // I ShapeText, EFTER fontIdMap-loopen (ersätt den nuvarande raden):
                Debug.WriteLine($"Shape: {tf.Font.Family}/{tf.Font.SubFamily} " +
                                $"usedFonts=[{string.Join(", ", usedFonts.Select(f => f.GetEnglishFontFamilyName()))}] " +
                                $"labels=[{string.Join(",", fontIdMap.Values)}]");
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

        // Pass 2: shape text using already-built providers from PdfDictionaries.ShapedProviders
        public static void ShapeText(PdfPageSettings pageSettings, PdfDictionaries dictionaries, PdfCellBase cell)
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
                var key = dictionaries.ResolveFontKey(pageSettings, tf.Font.Family, tf.Font.SubFamily);
                if (!dictionaries.ShapedProviders.TryGetValue(key, out var provider))
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
                var fontIdMap = new Dictionary<byte, string>();
                for (byte fontId = 0; fontId < usedFonts.Count; fontId++)
                {
                    var font = usedFonts[fontId];
                    var loadedKey = new FontKey(font.GetEnglishFontFamilyName(), font.NameTable.GetSubfamilyEnum());
                    if (!dictionaries.Fonts.ContainsKey(loadedKey))
                    {
                        int label = dictionaries.Fonts.Count > 0
                            ? dictionaries.Fonts.Last().Value.labelNumber + 1
                            : 1;
                        var fontResource = new PdfFontResource(font.GetEnglishFontFamilyName(), font.NameTable.GetSubfamilyEnum(), label, pageSettings);
                        fontResource.fontData = font;
                        dictionaries.Fonts.Add(loadedKey, fontResource);
                    }
                    fontIdMap[fontId] = dictionaries.Fonts[loadedKey].Label;
                }
                Debug.WriteLine($"Shape: {tf.Font.Family}/{tf.Font.SubFamily} " +
                $"usedFonts=[{string.Join(", ", usedFonts.Select(f => f.GetEnglishFontFamilyName()))}] " +
                $"labels=[{string.Join(",", fontIdMap.Values)}]");

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
                double wrapWidth = cell.Width;
                cell.TextLines = cell.ContentAligmnet.WrapText
                    ? cell.TextLayoutEngine.WrapRichTextLineCollection(cell.TextFragments, wrapWidth)
                    : cell.TextLayoutEngine.WrapRichTextLineCollection(cell.TextFragments, double.MaxValue);
            }
            cell.TotalTextLength = totalTextLength;
        }
    }
}