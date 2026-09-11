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
                var options = BuildShapingOptions(pageSettings);
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
                double wrapWidth = (cell.Merged && cell.Main == null) ? cell.Width : cell.ColumnWidth;
                cell.TextLines = cell.ContentAligmnet?.IsVertical == true
                    ? cell.TextLayoutEngine.BuildVerticalLineCollection(cell.TextFragments)
                    : cell.ContentAligmnet.WrapText
                        ? cell.TextLayoutEngine.WrapRichTextLineCollection(cell.TextFragments, wrapWidth)
                        : cell.TextLayoutEngine.WrapRichTextLineCollection(cell.TextFragments, double.MaxValue);
            }
            cell.TotalTextLength = totalTextLength;
        }

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
                var options = BuildShapingOptions(pageSettings);
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

        /// <summary>
        /// Builds the ShapingOptions used for one text fragment, from the caller's requested
        /// GsubFeature/GposFeature flags.
        /// </summary>
        /// <remarks>
        /// GsubFeature.None / GposFeature.None need special handling here rather than a plain
        /// pass-through of ToTagList's empty list. TextShaper.ApplyPositioning treats an empty or
        /// null GposFeatures list as "apply every GPOS feature" for kerning and mark positioning
        /// (though NOT for single adjustment, which treats it as "apply nothing" - the two
        /// disagree on empty/null already, independently of this method). That documented
        /// contract has other, unrelated callers (measurement, rich text default, benchmarks) and
        /// is not changed here. Instead, None is handled at the source: when the caller asks for
        /// no GPOS/GSUB features at all, ApplyPositioning/ApplySubstitutions are turned off
        /// outright, which is unambiguous regardless of what an empty tag list would otherwise be
        /// interpreted as further down.
        /// </remarks>
        internal static ShapingOptions BuildShapingOptions(PdfPageSettings pageSettings)
        {
            var options = ShapingOptions.Default;

            options.GsubFeatures = GsubFeatureTags.ToTagList(pageSettings.GsubFeatures);
            options.GposFeatures = GposFeatureTags.ToTagList(pageSettings.GposFeatures);

            options.ApplySubstitutions = pageSettings.GsubFeatures != GsubFeature.None;
            options.ApplyPositioning = pageSettings.GposFeatures != GposFeature.None;

            return options;
        }
    }
}