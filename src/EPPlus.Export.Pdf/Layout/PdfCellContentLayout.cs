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
        public PdfCellAlignmentData CellAlignmentData;
        public bool Clip;
        public Rect Clipping;
        public bool IsHeaderFooter;
        public bool IsHeading;
        public bool IsPrintTitle;
        public TextLayoutEngine textLayoutEngine;
        public TextLineCollection TextLines;
        public double LeftTextSpillLength = 0d;
        public double RightTextSpillLength = 0d;
        private double bottomMargin = 3.5d; //Guessed number
        private double rightMargin = 1.4d; //I guessed this one too..

        public List<PdfShapedText> ShapedTexts { get; set; }

        public PdfCellContentLayout(PdfPageSettings pageSettings, PdfDictionaries dictionaries, PdfCellBase cell, MergedCellDrawInfo mergedCellInfo, double x, double y, double width, double height, double scaleX = 1, double scaleY = 1, double rotation = 0, Transform parent = null)
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
            LocalPosition = CalculateAlignment(cell.Text, TextLines.LineFragments[0].Width, totalTextHeight, firstLineAscent, lastLineAscent, LocalPosition.X, LocalPosition.Y, cell.Width, height);
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
            var newX = CalculateHorizontalAlignment(TextLines.LineFragments[0].OriginalTextFragment.Text, TextLines[0].Width, LocalPosition.X, width, 0);
            LocalPosition = new Vector2 (newX, LocalPosition.Y);
        }

        private double CalculateVerticalAlignment(string text, double textHeight, double firstAscent, double lastAscent, double y, double height, double padding)
        {
            double newY = y;
            switch (CellAlignmentData.VerticalAlignment)
            {
                case ExcelVerticalAlignment.Top:
                case ExcelVerticalAlignment.Distributed:
                case ExcelVerticalAlignment.Justify:
                    newY = (y + height) - padding - firstAscent;
                    break;
                case ExcelVerticalAlignment.Center:
                    newY = y + (height + textHeight - firstAscent - lastAscent) / 2d;
                    break;
                case ExcelVerticalAlignment.Bottom:
                    newY = y + padding + textHeight - lastAscent;
                    break;
            }
            return newY;
        }

        private double CalculateHorizontalAlignment(string text, double textLength, double x, double width, double padding)
        {
            double newX = x;
            switch (CellAlignmentData.HorizontalAlignment)
            {
                case ExcelHorizontalAlignment.Fill:
                case ExcelHorizontalAlignment.General:
                    if (double.TryParse(text, out double value))
                    {
                        newX = x + (width - textLength) - padding;
                    }
                    else
                    {
                        newX = x + padding;
                    }
                    break;
                case ExcelHorizontalAlignment.Left:
                case ExcelHorizontalAlignment.Justify:
                case ExcelHorizontalAlignment.Distributed:
                    newX = x + padding;
                    break;
                case ExcelHorizontalAlignment.Center:
                case ExcelHorizontalAlignment.CenterContinuous:
                    newX = x + (width - textLength) / 2d;
                    break;
                case ExcelHorizontalAlignment.Right:
                    newX = x + (width - textLength) - padding;
                    break;
            }
            return newX;
        }

        private Vector2 CalculatePositionFromRotation(double textLength, double x, double y)
        {
            double newX = x;
            double newY = y;
            if (CellAlignmentData.TextRotation < 0)
            {
                double rot = CellAlignmentData.TextRotation * System.Math.PI / 180.0;
                newX += textLength * (1 - System.Math.Cos(rot));
                newY -= textLength * System.Math.Sin(rot);
            }
            else if (CellAlignmentData.TextRotation > 0)
            {
                double rot = CellAlignmentData.TextRotation * System.Math.PI / 180.0;
                newX += textLength * (1 - System.Math.Cos(rot));
            }
            return new Vector2(newX, newY);
        }

        private Vector2 CalculateAlignment(string text, double textLength, double textHeight, double firstLineAscent, double lastLineAscent, double x, double y, double width, double height)
        {
            if (CellAlignmentData.TextRotation != 0 && !CellAlignmentData.IsVertical)
            {
                return CalculateRotatedAlignment(textLength, textHeight, firstLineAscent, x, y, width, height);
            }
            double newX = CalculateHorizontalAlignment(text, textLength, x, width, rightMargin);
            double newY = CalculateVerticalAlignment(text, textHeight, firstLineAscent, lastLineAscent, y, height, 0d);
            return CalculatePositionFromRotation(textLength, newX, newY);
        }

        private Vector2 CalculateRotatedAlignment(double textLength, double textHeight, double firstLineAscent, double x, double y, double width, double height)
        {
            double ascent = firstLineAscent;
            double descent = textHeight - firstLineAscent;
            if (descent < 0d) descent = 0d;
            double theta = CellAlignmentData.TextRotation * System.Math.PI / 180.0;
            double cos = System.Math.Cos(theta);
            double sin = System.Math.Sin(theta);
            // Bounding box of the rotated block. Reading direction spans [0, textLength];
            // the cross (line-height) direction spans [-descent, ascent].
            double[] gx = { 0d, textLength, 0d, textLength };
            double[] gy = { ascent, ascent, -descent, -descent };
            double minX = double.MaxValue, maxX = double.MinValue, minY = double.MaxValue, maxY = double.MinValue;
            for (int i = 0; i < 4; i++)
            {
                double ux = cos * gx[i] - sin * gy[i];
                double uy = sin * gx[i] + cos * gy[i];
                if (ux < minX) minX = ux;
                if (ux > maxX) maxX = ux;
                if (uy < minY) minY = uy;
                if (uy > maxY) maxY = uy;
            }
            double blockWidth = maxX - minX;
            double blockHeight = maxY - minY;
            double bx;
            switch (CellAlignmentData.HorizontalAlignment)
            {
                case ExcelHorizontalAlignment.Left:
                    bx = x + rightMargin; break;
                case ExcelHorizontalAlignment.Right:
                    bx = x + width - blockWidth - rightMargin; break;
                default: // Center / General
                    bx = x + (width - blockWidth) / 2d; break;
            }
            double by;
            switch (CellAlignmentData.VerticalAlignment)
            {
                case ExcelVerticalAlignment.Top:
                    by = y + height - blockHeight - bottomMargin; break;
                case ExcelVerticalAlignment.Bottom:
                    by = y + bottomMargin; break;
                default: // Center / Justify / Distributed
                    by = y + (height - blockHeight) / 2d; break;
            }
            // Convert the bounding-box lower-left back to the baseline origin the matrix expects.
            return new Vector2(bx - minX, by - minY);
        }

        // Set clipping to the cell's own bounds. cellY is the top edge (same convention as the constructor).
        internal void SetupClipping(double cellX, double cellY, double cellWidth, double cellHeight)
        {
            Clip = true;
            Clipping = new Rect()
            {
                X = cellX + rightMargin,
                Y = cellY - cellHeight,   // bottom-left corner in PDF space
                Width = cellWidth - rightMargin * 2,
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
