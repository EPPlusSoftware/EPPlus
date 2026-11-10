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
using EPPlus.Export.Pdf.Enums;
using EPPlus.Export.Pdf.Resources;
using EPPlus.Export.Pdf.Settings;
using EPPlus.Fonts.OpenType.Integration;
using EPPlus.Graphics;
using OfficeOpenXml.Interfaces.Fonts;
using System.Collections.Generic;
using System.Diagnostics;
using Vector2 = EPPlus.Graphics.Geometry.Vector2;

namespace EPPlus.Export.Pdf.Layout
{
    [DebuggerDisplay("Content: {Name}")]
    internal class PdfCellContentLayout : Transform
    {
        internal const double CellHorizontalInset = 1.8d;  //I guessed this number

        public PdfCellAlignmentData CellAlignmentData;
        public bool Clip;
        public Rect Clipping;
        public bool IsHeaderFooter;
        public bool IsAddedContent;
        public bool IsHeading;
        public bool IsPrintTitle;
        public TextLayoutEngine textLayoutEngine;
        public TextLineCollection TextLines;
        public double LeftTextSpillLength = 0d;
        public double RightTextSpillLength = 0d;
        public double ScaleAnchorX;
        public double ScaleAnchorY;
        private double bottomMargin = 3.5d; //Guessed number
        private double rightMargin = CellHorizontalInset;

        public List<PdfShapedText> ShapedTexts { get; set; }

        public PdfCellContentLayout(PdfPageSettings pageSettings, PdfDictionaries dictionaries, PdfCellBase cell, double x, double y, double width, double height, double scaleX = 1, double scaleY = 1, double rotation = 0, Transform parent = null)
            : base(x, y-height, width, height, scaleX, scaleY, rotation, parent)
        {
            Z = 2;
            CellAlignmentData = cell.ContentAligmnet;
            TextLines = cell.TextLines;
            ShapedTexts = cell.ShapedTexts;
            textLayoutEngine = cell.TextLayoutEngine;
            double totalTextHeight = 0d;
            foreach (var line in TextLines)
            {
                totalTextHeight += line.LargestAscent + line.LargestDescent;
            }
            double firstLineAscent = TextLines[0].LargestAscent;
            double lastLineAscent = TextLines[TextLines.Count - 1].LargestAscent;
            LocalPosition = PdfCellAlignment.CalculateAlignment(CellAlignmentData, cell.Text, TextLines.LineFragments[0].Width, totalTextHeight, firstLineAscent, lastLineAscent, LocalPosition.X, LocalPosition.Y, width, height, rightMargin, bottomMargin);
        }

        public PdfCellContentLayout(PdfPageSettings pageSettings, PdfDictionaries dictionaries, PdfHeaderFooter headerFooter, double x, double y, double width, double height, double scaleX = 1, double scaleY = 1, double rotation = 0, Transform parent = null)
            : base(x, y, width, height, scaleX, scaleY, rotation, parent)
        {
            Z = 2;
            TextLines = headerFooter.Content.TextLines;
            ShapedTexts = headerFooter.Content.ShapedTexts;
            textLayoutEngine = headerFooter.Content.TextLayoutEngine;
            CellAlignmentData = headerFooter.Content.ContentAligmnet;
            double totalTextHeight = 0d;
            foreach (var line in TextLines)
            {
                totalTextHeight += line.LargestAscent + line.LargestDescent;
            }
            var newX = PdfCellAlignment.CalculateHorizontalAlignment(CellAlignmentData, TextLines.LineFragments[0].OriginalTextFragment.Text, TextLines[0].Width, LocalPosition.X, width, 0);
            LocalPosition = new Vector2 (newX, LocalPosition.Y);
        }

        internal void SetupClipping(double cellX, double cellY, double cellWidth, double cellHeight, bool spillLeft = false, bool spillRight = false)
        {
            double li = spillLeft ? 0d : rightMargin;
            double ri = spillRight ? 0d : rightMargin;
            Clip = true;
            Clipping = new Rect()
            {
                X = cellX + li,
                Y = cellY - cellHeight,
                Width = cellWidth - li - ri,
                Height = cellHeight
            };
        }

        internal void GidsAndCharMap(PdfDictionaries dictionaries)
        {
            foreach (var tf in ShapedTexts)
            {
                var usedFonts = tf.UsedFonts;

                foreach (var glyph in tf.ShapedText.Glyphs)
                {
                    if (glyph.FontId >= usedFonts.Count)
                        continue;

                    var font = usedFonts[glyph.FontId];
                    var key = new FontKey(font.GetEnglishFontFamilyName(), font.NameTable.GetSubfamilyEnum());

                    dictionaries.Fonts[key].Gids.Add(glyph.GlyphId);
                    dictionaries.Fonts[key].fontData = font;

                    if (!dictionaries.Fonts[key].charactermappings.ContainsKey(glyph.GlyphId))
                    {
                        var chars = ExtractCharactersForGlyph(glyph, tf.ShapedText.OriginalText);
                        if (!string.IsNullOrEmpty(chars))
                        {
                            dictionaries.Fonts[key].charactermappings[glyph.GlyphId] = chars;
                        }
                    }
                }
            }
        }
        private string ExtractCharactersForGlyph(ShapedGlyph glyph, string textLine)
        {
            var chars = new System.Text.StringBuilder();
            for (int i = 0; i < glyph.CharCount && glyph.ClusterIndex + i < textLine.Length; i++)
            {
                chars.Append(textLine[glyph.ClusterIndex + i]);
            }
            return chars.ToString();
        }
    }
}
