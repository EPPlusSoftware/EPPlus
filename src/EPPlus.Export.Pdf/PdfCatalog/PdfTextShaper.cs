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

        public static void LayoutAndShapeText(PdfPageSettings pageSettings, PdfDictionaries dictionaries, PdfCell Cell)
        {
            var totalTextLength = 0d;
            var maxLineHeight = 0d;
            if (Cell.TextFragments == null) return;
            Cell.ShapedTexts = new List<PdfShapedText>();
            for (int i = 0; i < Cell.TextFragments.Count; i++)
            {
                var tf = Cell.TextFragments[i];
                Cell.ShapedTexts.Add(new PdfShapedText());
                var st = Cell.ShapedTexts[i];
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

                if (Cell.ContentAligmnet.WrapText)
                    Cell.TextLines = layoutEngine.WrapRichTextLineCollection(Cell.TextFragments, Cell.ColumnWidth);
                else
                    Cell.TextLines = layoutEngine.WrapRichTextLineCollection(Cell.TextFragments, double.MaxValue);

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

                Cell.TextLayoutEngine = layoutEngine;
                st.ShapedText = shaped;
                var textWdith = st.ShapedText.GetWidthInPoints((float)tf.Font.Size);
                var textHeight = st.ShapedText.GetLineHeightInPoints((float)tf.Font.Size);
                //tf.TextLength = textWdith;
                //tf.TextHeight = textHeight;
                totalTextLength += textWdith;
                maxLineHeight = Math.Max(textHeight, maxLineHeight);
                st.FontIdMap = fontIdMap;
                st.UsedFonts = usedFonts;
                Cell.TextFragments[i] = tf;
                Cell.ShapedTexts[i] = st;
            }
            Cell.TotalTextLength = totalTextLength;
            //if (Cell.ContentAligmnet.WrapText)
            //{
            //    var textFragments = GetTextFragments(Cell.TextFormats);
            //    var wrappedLines = Cell.TextLayoutEngine.WrapRichTextLines(textFragments, Cell.Width);

            //}
        }

        private static List<TextFragment> GetTextFragments(List<TextFragment> textFragments)
        {
            var fragments = new List<TextFragment>(textFragments.Count);

            foreach (var tf in textFragments)
            {
                var fragment = new TextFragment
                {
                    Text = tf.Text,
                    Font = new MeasurementFont
                    {
                        FontFamily = tf.Font.FontFamily,
                        Size = (float)tf.Font.Size,
                        Style = GetMeasurementFontStyle(tf)
                    }
                };
                fragments.Add(fragment);
            }

            return fragments;
        }

        private static MeasurementFontStyles GetMeasurementFontStyle(TextFragment tf)
        {
            var style = (tf.RichTextOptions.Bold ? MeasurementFontStyles.Bold : 0)
                      | (tf.RichTextOptions.Italic ? MeasurementFontStyles.Italic : 0)
                      | (tf.RichTextOptions.StrikeType > 1 ? MeasurementFontStyles.Strikeout : 0)
                      | (tf.RichTextOptions.UnderlineType != 12 ? MeasurementFontStyles.Underline : 0);

            return style == 0 ? MeasurementFontStyles.Regular : (MeasurementFontStyles)style;
        }
    }
}
