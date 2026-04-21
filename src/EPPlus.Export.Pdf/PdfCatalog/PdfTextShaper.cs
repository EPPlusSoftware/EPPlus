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
            if (Cell.TextFormats == null) return;
            for (int i = 0; i < Cell.TextFormats.Count; i++)
            {
                var fd = Cell.TextFormats[i];
                fd.FontProvider = dictionaries.Fonts[fd.FullFontName].fontSubsetManager.CreateSubsettedProvider();

                if (!shaperCache.TryGetValue(fd.FontProvider, out var shaper))
                {
                    shaper = new TextShaper(fd.FontProvider);
                    shaperCache[fd.FontProvider] = shaper;
                }

                if (!layoutEngineCache.TryGetValue(fd.FontProvider, out var layoutEngine))
                {
                    layoutEngine = new TextLayoutEngine(shaper);
                    layoutEngineCache[fd.FontProvider] = layoutEngine;
                }

                var options = ShapingOptions.Default;
                options.ApplyPositioning = true;
                options.ApplySubstitutions = true;

                var shaped = shaper.Shape(fd.Text, options);
                var usedFonts = shaper.GetUsedFonts().ToList();
                var fontIdMap = new Dictionary<byte, string>();

                var allProviderFonts = fd.FontProvider.GetAllFonts().ToList();

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
                fd.ShapedText = shaped;
                var textWdith = fd.ShapedText.GetWidthInPoints((float)fd.FontSize);
                var textHeight = fd.ShapedText.GetLineHeightInPoints((float)fd.FontSize);
                fd.TextLength = textWdith;
                fd.TextHeight = textHeight;
                totalTextLength += textWdith;
                maxLineHeight = Math.Max(textHeight, maxLineHeight);
                fd.FontIdMap = fontIdMap;
                fd.UsedFonts = usedFonts;
                Cell.TextFormats[i] = fd;
            }
            Cell.TotalTextLength = totalTextLength;
            //if (Cell.ContentAligmnet.WrapText)
            //{
            //    var textFragments = GetTextFragments(Cell.TextFormats);
            //    var wrappedLines = Cell.TextLayoutEngine.WrapRichTextLines(textFragments, Cell.Width);

            //}
        }

        private static List<TextFragment> GetTextFragments(List<PdfTextFormat> textFormats)
        {
            var fragments = new List<TextFragment>(textFormats.Count);

            foreach (var tf in textFormats)
            {
                var fragment = new TextFragment
                {
                    Text = tf.Text,
                    Font = new MeasurementFont
                    {
                        FontFamily = tf.FontName,
                        Size = (float)tf.FontSize,
                        Style = GetMeasurementFontStyle(tf)
                    }
                };
                fragments.Add(fragment);
            }

            return fragments;
        }

        private static MeasurementFontStyles GetMeasurementFontStyle(PdfTextFormat tf)
        {
            var style = (tf.Bold ? MeasurementFontStyles.Bold : 0)
                      | (tf.Italic ? MeasurementFontStyles.Italic : 0)
                      | (tf.Strike ? MeasurementFontStyles.Strikeout : 0)
                      | (tf.Underline ? MeasurementFontStyles.Underline : 0);

            return style == 0 ? MeasurementFontStyles.Regular : (MeasurementFontStyles)style;
        }
    }
}
