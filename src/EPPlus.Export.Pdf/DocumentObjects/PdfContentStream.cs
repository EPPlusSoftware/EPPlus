/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  10/07/2025         EPPlus Software AB           EPPlus.Fonts.OpenType 1.0
 *************************************************************************************************/
using EPPlus.Graphics;
using EPPlus.Graphics.Geometry;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using EPPlus.Fonts.OpenType.Integration;
using EPPlus.Export.Pdf.Layout;
using EPPlus.Export.Pdf.Helpers;
using EPPlus.Export.Pdf.Resources;
using EPPlus.Export.Pdf.Settings;
using EPPlus.Export.Pdf.Enums;

namespace EPPlus.Export.Pdf.DocumentObjects
{
    internal class PdfContentStream : PdfObject
    {
        private readonly List<string> commands = new List<string>();

        public PdfContentStream(int objectNumber, string command = null, int version = 0)
            : base(objectNumber, version)
        {
            if (!string.IsNullOrEmpty(command))
            {
                commands.Add(command);
            }
        }

        public void AddCommand(string command)
        {
            commands.Add(command);
        }

        public void AddCellLayout(PdfCellLayout cell, string label)
        {
            if (cell.Size.X <= 0d || cell.Size.Y <= 0d) return;
            if (cell.CellFillData.GradientFillData != null && cell.CellFillData.PatternStyle != ExcelFillStyle.Solid)
            {
                commands.Add($"% Pattern Start: {cell.Name}");
                commands.Add("q");
                commands.Add("/Pattern cs");
                commands.Add($"/{label} scn");
                commands.Add($"{cell.LocalPosition.X.ToPdfString()} {cell.LocalPosition.Y.ToPdfString()} {cell.Size.X.ToPdfString()} {cell.Size.Y.ToPdfString()} re");
                commands.Add("f");
                commands.Add("Q");
                commands.Add($"% Pattern End: {cell.Name}");
            }
            else if (cell.CellFillData.BackgroundColor != Color.Empty && cell.CellFillData.PatternStyle == ExcelFillStyle.Solid)
            {
                commands.Add($"% Solid Fill Start: {cell.Name}");
                commands.Add("q");
                commands.Add($"{GridLine.HalfWidth.ToPdfString()} w");
                commands.Add(cell.CellFillData.BackgroundColor.ToFillCommand());
                commands.Add(cell.CellFillData.enhanceGridLine ? Color.Black.ToStrokeCommand() : cell.CellFillData.BackgroundColor.ToStrokeCommand());
                commands.Add($"{cell.LocalPosition.X.ToPdfString()} {cell.LocalPosition.Y.ToPdfString()} {cell.Size.X.ToPdfString()} {cell.Size.Y.ToPdfString()} re");
                commands.Add("B");
                commands.Add("Q");
                commands.Add($"% Solid Fill End: {cell.Name}");
            }
            else if (cell.CellFillData.PatternStyle != ExcelFillStyle.None)
            {
                commands.Add($"% Pattern Start: {cell.Name}");
                // Draw the solid cell background only when one is set. The pattern
                // tile already fills itself with its own background color, so the
                // pattern must be rendered regardless of whether the cell has a
                // separate background fill (it may be Color.Empty).
                if (cell.CellFillData.BackgroundColor != Color.Empty)
                {
                    commands.Add("q");
                    commands.Add($"{GridLine.HalfWidth.ToPdfString()} w");
                    commands.Add(cell.CellFillData.BackgroundColor.ToFillCommand());
                    commands.Add(cell.CellFillData.BackgroundColor.ToStrokeCommand());
                    commands.Add($"{cell.LocalPosition.X.ToPdfString()} {cell.LocalPosition.Y.ToPdfString()} {cell.Size.X.ToPdfString()} {cell.Size.Y.ToPdfString()} re");
                    commands.Add("B");
                    commands.Add("Q");
                }

                commands.Add("q");
                commands.Add("/Pattern cs");
                commands.Add($"/{label} scn");
                commands.Add($"{cell.LocalPosition.X.ToPdfString()} {cell.LocalPosition.Y.ToPdfString()} {cell.Size.X.ToPdfString()} {cell.Size.Y.ToPdfString()} re");
                commands.Add("f");
                commands.Add("Q");
                commands.Add($"% Pattern End: {cell.Name}");
            }
        }

        public void AddBorderLayout(PdfCellBorderLayout cell)
        {
            var borderRenderer = new PdfBorderRenderer(cell);
            borderRenderer.RenderBorder(this);
        }

        public void AddText(PdfCellContentLayout cell, Vector2 position, double textRotation, PdfDictionaries dictionaries, PdfPageSettings pageSettings)
        {
            double advanceY = 0d;
            double line0Width = cell.TextLines.Count > 0 ? cell.TextLines[0].Width : 0d;
            double rotation = textRotation * System.Math.PI / 180.0;
            for (int k = 0; k < cell.TextLines.Count; k++)
            {
                var line = cell.TextLines[k];
                double lineOffsetX = 0d;
                switch (cell.CellAlignmentData.HorizontalAlignment)
                {
                    case ExcelHorizontalAlignment.Right:
                        lineOffsetX = line0Width - line.Width;
                        break;
                    case ExcelHorizontalAlignment.Center:
                        lineOffsetX = (line0Width - line.Width) / 2d;
                        break;
                }
                double advanceX = 0;
                for (int i = 0; i < line.LineFragments.Count; i++)
                {
                    var textFormat = line.LineFragments[i];
                    // find which ShapedText owns this fragment
                    int shapedTextIndex = 0;
                    int shapedTextCharStart = 0;
                    for (int s = 0; s < cell.ShapedTexts.Count; s++)
                    {
                        int shapedTextCharCount = cell.ShapedTexts[s].ShapedText.Glyphs.Sum(g => g.CharCount);
                        if (textFormat.StartFullTextIdx < shapedTextCharStart + shapedTextCharCount)
                        {
                            shapedTextIndex = s;
                            break;
                        }
                        shapedTextCharStart += shapedTextCharCount;
                    }
                    var shapedText = cell.ShapedTexts[shapedTextIndex];

                    // find the starting glyph within that ShapedText using StartRtIdx
                    int glyphStart = 0;
                    int charOffsetInShapedText = textFormat.StartFullTextIdx - shapedTextCharStart;
                    int rtCharCount = 0;
                    for (int g = 0; g < shapedText.ShapedText.Glyphs.Length; g++)
                    {
                        if (rtCharCount >= charOffsetInShapedText)
                        {
                            glyphStart = g;
                            break;
                        }
                        rtCharCount += shapedText.ShapedText.Glyphs[g].CharCount;
                    }
                    var originalFragment = ((TextFragment)textFormat.OriginalTextFragment);
                    while (glyphStart < shapedText.ShapedText.Glyphs.Length && shapedText.ShapedText.Glyphs[glyphStart].GlyphId == 0)
                    {
                        shapedTextIndex++;
                        if (shapedTextIndex >= cell.ShapedTexts.Count)
                        {
                            break; // temp safety
                        }
                        shapedText = cell.ShapedTexts[shapedTextIndex];
                        glyphStart = 0;
                    }

                    var richInfo = originalFragment.RichTextOptions;
                    var textLength = shapedText.ShapedText.GetWidthInPoints((float)richInfo.Size);
                    var color = richInfo.FontColor;
                    var fontResource = dictionaries.GetFont(pageSettings, richInfo.Family, richInfo.SubFamily);
                    double size = richInfo.Size;
                    double scale = textFormat.OriginalTextFragment.RichTextOptions.Size / fontResource.fontData.HeadTable.UnitsPerEm;
                    Matrix3x3 textMatrix = new Matrix3x3(System.Math.Cos(rotation), System.Math.Sin(rotation), -System.Math.Sin(rotation), System.Math.Cos(rotation), position.X + lineOffsetX, position.Y + advanceY);
                    commands.Add("BT");
                    textMatrix = textMatrix * Matrix3x3.Translation(advanceX, 0);
                    if (richInfo.SuperScript)
                    {
                        var supOffX = fontResource.fontData.Os2Table.ySuperscriptXOffset * scale;
                        var supOffY = fontResource.fontData.Os2Table.ySuperscriptYOffset * scale;
                        var supSizeY = fontResource.fontData.Os2Table.ySuperscriptYSize * scale;
                        textMatrix = textMatrix * Matrix3x3.Translation(supOffX, supOffY);
                        size = supSizeY;
                    }
                    else if (richInfo.SubScript)
                    {
                        var supOffX = fontResource.fontData.Os2Table.ySubscriptXOffset * scale;
                        var supOffY = fontResource.fontData.Os2Table.ySubscriptYOffset * scale;
                        var supSizeY = fontResource.fontData.Os2Table.ySubscriptYSize * scale;
                        textMatrix = textMatrix * Matrix3x3.Translation(supOffX, supOffY);
                        size = supSizeY;
                    }
                    if (richInfo.UnderlineType != 12)
                    {
                        var underlinePos = fontResource.fontData.PostTable.underlinePosition * scale;
                        var underlineWidth = fontResource.fontData.PostTable.underlineThickness * scale;
                        var start = textMatrix.Transform(new Vector2(0, underlinePos));
                        var end = textMatrix.Transform(new Vector2(textLength, underlinePos));
                        commands.Add($"{underlineWidth.ToPdfString()} w");
                        commands.Add($"{start.X.ToPdfString()} {start.Y.ToPdfString()} m");
                        commands.Add($"{end.X.ToPdfString()} {end.Y.ToPdfString()} l");
                        commands.Add($"S");
                    }
                    if (richInfo.StrikeType > 1)
                    {
                        var strikePos = fontResource.fontData.Os2Table.yStrikeoutPosition * scale;
                        var strikeWidth = fontResource.fontData.Os2Table.yStrikeoutSize * scale;
                        var start = textMatrix.Transform(new Vector2(0, strikePos));
                        var end = textMatrix.Transform(new Vector2(textLength, strikePos));
                        commands.Add($"{strikeWidth.ToPdfString()} w");
                        commands.Add($"{start.X.ToPdfString()} {start.Y.ToPdfString()} m");
                        commands.Add($"{end.X.ToPdfString()} {end.Y.ToPdfString()} l");
                        commands.Add($"S");
                    }
                    commands.Add(color.ToFillCommand());
                    commands.Add($"{textMatrix.A.ToPdfStringF4()} {textMatrix.B.ToPdfStringF4()} {textMatrix.C.ToPdfStringF4()} {textMatrix.D.ToPdfStringF4()} {textMatrix.E.ToPdfStringF4()} {textMatrix.F.ToPdfStringF4()} Tm");

                    // FIX: Always use fontIdMap to determine the initial font.
                    // FontId=0 does NOT always mean "primary font" — when the text starts
                    // with a fallback character (e.g. emoji), FontId=0 IS the fallback font.
                    // The fontIdMap correctly maps FontId → PDF font label in all cases.
                    byte currentFontId = shapedText.ShapedText.Glyphs.Length > 0 ? shapedText.ShapedText.Glyphs[0].FontId : (byte)0;
                    string currentFontLabel = shapedText.FontIdMap.ContainsKey(currentFontId)
                        ? shapedText.FontIdMap[currentFontId]
                        : fontResource.Label;
                    commands.Add($"/{currentFontLabel} {size.ToPdfString()} Tf");
                    int fragmentCharCount = textFormat.Text.Length;
                    int charsRendered = 0;
                    var sb = new StringBuilder();
                    sb.Append("[");
                    for (int j = glyphStart; j < shapedText.ShapedText.Glyphs.Length; j++)
                    {
                        if (charsRendered >= fragmentCharCount)
                            break;
                        var glyph = shapedText.ShapedText.Glyphs[j];
                        if (glyph.FontId != currentFontId)
                        {
                            // Close TJ array, switch font, open new TJ array
                            sb.Append("] TJ\n");
                            sb.Append($"/{shapedText.FontIdMap[glyph.FontId]} {size.ToPdfString()} Tf\n");
                            sb.Append("[");
                            currentFontId = glyph.FontId;
                        }
                        sb.Append($"<{glyph.GlyphId:X4}>");
                        int kerning = glyph.XAdvance - glyph.BaseAdvance;
                        if (kerning != 0)
                        {
                            double adjustment = -(kerning * 1000.0 / 1000);
                            sb.Append($" {adjustment.ToPdfStringF0()}");
                        }
                        if (j < shapedText.ShapedText.Glyphs.Length - 1)
                        {
                            sb.Append(" ");
                        }
                        charsRendered += glyph.CharCount;
                    }
                    advanceX += textLength;
                    commands.Add(sb.ToString() + "] TJ");
                    commands.Add("ET");
                }
                advanceY -= (line.LargestAscent + line.LargestDescent);
            }
        }

        public void AddCellContentLayout(PdfCellContentLayout cell, PdfDictionaries dictionaries, PdfPageSettings pageSettings)
        {
            commands.Add($"% Content Start: {cell.Name}");
            commands.Add("q");
            if (cell.Clip) AddClipping(cell);
            AddText(cell, cell.LocalPosition, cell.CellAlignmentData.TextRotation, dictionaries, pageSettings);
            commands.Add("Q");
            commands.Add($"% Content End: {cell.Name}");
        }

        private void AddClipping(PdfCellContentLayout cell)
        {
            commands.Add($"{cell.Clipping.X.ToPdfString()} {cell.Clipping.Y.ToPdfString()} {cell.Clipping.Width.ToPdfString()} {cell.Clipping.Height.ToPdfString()} re W n");
        }

        public void AddInnerGridLines(Transform pageLayout)
        {
            if (pageLayout is not PdfPageLayout pl) return;
            if (pl.isCommentsPage) return;

            commands.Add($"% Gridlines Start");
            commands.Add("q");
            commands.Add($"{GridLine.Width.ToPdfString()} w");
            commands.Add(Color.Black.ToFillCommand());
            foreach (var line in pl.GridLines)
            {
                string w, h;
                if (line.X1 == line.X2)
                {
                    w = GridLine.Width.ToPdfStringF4();
                    h = System.Math.Abs(line.Y2 - line.Y1).ToPdfStringF4();
                }
                else
                {
                    w = System.Math.Abs(line.X2 - line.X1).ToPdfStringF4();
                    h = GridLine.Width.ToPdfStringF4();
                }
                var x = Math.Min(line.X1, line.X2);
                var y = Math.Min(line.Y1, line.Y2);
                commands.Add($"{x.ToPdfStringF4()} {y.ToPdfStringF4()} {w} {h} re");
            }
            commands.Add("f");
            commands.Add("Q");
            commands.Add($"% Gridlines End");
        }

        public void AddPrintTitleGridLines(Transform pageLayout)
        {
            if (pageLayout is not PdfPageLayout pl) return;
            if (pl.isCommentsPage) return;

            commands.Add($"% Print Title Gridlines Start");
            commands.Add("q");
            commands.Add($"{GridLine.Width.ToPdfString()} w");
            commands.Add(Color.Black.ToFillCommand());
            foreach (var line in pl.PrintTitleGridLines)
            {
                string w, h;
                if (line.X1 == line.X2)
                {
                    w = GridLine.Width.ToPdfStringF4();
                    h = System.Math.Abs(line.Y2 - line.Y1).ToPdfStringF4();
                }
                else
                {
                    w = System.Math.Abs(line.X2 - line.X1).ToPdfStringF4();
                    h = GridLine.Width.ToPdfStringF4();
                }
                var x = Math.Min(line.X1, line.X2);
                var y = Math.Min(line.Y1, line.Y2);
                commands.Add($"{x.ToPdfStringF4()} {y.ToPdfStringF4()} {w} {h} re");
            }
            commands.Add("f");
            commands.Add("Q");
            commands.Add($"% Print Title Gridlines End");
        }

        public void AddOuterGridBorder(Transform pageLayout)
        {
            if (pageLayout is not PdfPageLayout pl) return;
            if (pl.isCommentsPage) return;

            commands.Add($"% Gridlines Border Start");
            commands.Add("q");
            commands.Add("1.0 w");
            commands.Add("2 J");
            commands.Add("[] 0 d");
            commands.Add(Color.Black.ToStrokeCommand());
            foreach (var line in pl.BorderLines)
            {
                commands.Add($"{line.X1.ToPdfStringF4()} {line.Y1.ToPdfStringF4()} m");
                commands.Add($"{line.X2.ToPdfStringF4()} {line.Y2.ToPdfStringF4()} l");
            }
            commands.Add("S");
            commands.Add("Q");
            commands.Add($"% Gridlines Border End");
        }

        public void AddMarginClipping(PdfPageLayout pageLayout, PdfPageSettings pageSettings)
        {
            if (pageLayout is not PdfPageLayout pl) return;
            if (pageLayout.isCommentsPage) return;
            commands.Add($"% Margin Clip Start");
            if (pl.BorderLines.Count == 0) return;
            // Derive the tight bounding box directly from BorderLines.
            // pageLayout is created with all-zero dimensions so ContentTop/Bottom/Left/Height
            // cannot be used here — they are always 0.
            double top = double.MinValue;
            double bottom = double.MaxValue;
            double left = double.MaxValue;
            double right = double.MinValue;
            foreach (var line in pl.BorderLines)
            {
                top = System.Math.Max(top, System.Math.Max(line.Y1, line.Y2));
                bottom = System.Math.Min(bottom, System.Math.Min(line.Y1, line.Y2));
                left = System.Math.Min(left, System.Math.Min(line.X1, line.X2));
                right = System.Math.Max(right, System.Math.Max(line.X1, line.X2));
            }
            right = System.Math.Min(right, pageSettings.PageSize.WidthPu);
            bottom = System.Math.Max(bottom, 0d);
            var pad = GridLine.Width * 4;
            var x = left + pl.HeadingWidth + pl.PrintTitleWidth - pad;
            var y = bottom - pad;
            var width = (right - left - pl.HeadingWidth - pl.PrintTitleWidth) + pad * 2;
            var height = (top - pl.HeadingHeight - pl.PrintTitleHeight - bottom) + pad * 2;
            commands.Add($"{x.ToPdfString()} {y.ToPdfString()} {width.ToPdfString()} {height.ToPdfString()} re W n");
        }

        internal override string RenderDictionary()
        {
            var content = string.Join("\n", commands.ToArray()) + "\n";
            var bytes = Encoding.ASCII.GetBytes(content);
            return $"<< /Length {bytes.Length} >>\n" + $"stream\n{content}endstream";
        }

        internal override void RenderDictionary(BinaryWriter bw)
        {
            var content = string.Join("\n", commands.ToArray()) + "\n";
            var body = PdfFlate.Compress(Encoding.ASCII.GetBytes(content));
            WriteAscii(bw, $"<< /Filter /FlateDecode /Length {body.Length} >>\nstream\n");
            bw.Write(body);
            WriteAscii(bw, "\nendstream");
        }
    }
}