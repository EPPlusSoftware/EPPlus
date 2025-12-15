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
using EPPlus.Graphics;
using EPPlus.Graphics.Math;
using System.Drawing;
using EPPlus.Export.Pdf.Pdfhelpers;
using EPPlus.Export.Pdf.PdfResources;
using EPPlus.Export.Pdf.PdfSettings;
using EPPlus.Fonts.OpenType;
using OfficeOpenXml;
using OfficeOpenXml.Interfaces.Drawing.Text;
using OfficeOpenXml.Style;
using System.Collections.Generic;
using System.Linq;
using System;

namespace EPPlus.Export.Pdf.PdfLayout
{
    internal class PdfCellContentLayout : Transform
    {
        public List<PdfCellTextLine> TextLines = new List<PdfCellTextLine>();
        public PdfCellAlignmentData CellAlignmentData;
        public bool Clip;
        public Rect Clipping;

        private double bottomMargin = 3.5d; //Guessed number
        private double rightMargin = 1.4d; //I guessed this one too..

        internal static FontMeasurerTrueType fontMeasurerTrueType = new FontMeasurerTrueType();
        internal static MeasurementFont font = new MeasurementFont();

        public PdfCellContentLayout(ExcelRangeBase cell, PdfPageSettings pageSettings, double x, double y, double width, double height, double scaleX = 1, double scaleY = 1, double rotation = 0, Transform parent = null, PdfDictionaries dictionaries = null)
            : base(x, y, width, height, scaleX, scaleY, rotation, parent)
        {
            this.cell = cell;
            if (cell.IsRichText)
            {
                HandleRichText(pageSettings, dictionaries, width, height, x, cell.Style.TextRotation);
                //HandleRichText(pageSettings, dictionaries, width, height);
            }
            else
            {
                HandleText(pageSettings, dictionaries, x, y, width, height, cell.Style.TextRotation);
            }
            CellAlignmentData = new PdfCellAlignmentData();
            CellAlignmentData.HorizontalAlignment = cell.Style.HorizontalAlignment;
            CellAlignmentData.VerticalAlignment = cell.Style.VerticalAlignment;
            CellAlignmentData.Indent = cell.Style.Indent;
            CellAlignmentData.WrapText = cell.Style.WrapText;
            CellAlignmentData.ShrinkToFit = cell.Style.ShrinkToFit; //Need to fix Transform issues and then implement a method that sets scale on the text object.
            CellAlignmentData.TextRotation = (cell.Style.TextRotation >= 90) ? ((cell.Style.TextRotation == 255) ? 0 : 90 - cell.Style.TextRotation) : cell.Style.TextRotation;
            CellAlignmentData.IsVertical = cell.Style.TextRotation == 255 ? true : false;
            CellAlignmentData.TextDirection = cell.Style.ReadingOrder;
            //LocalPosition = CalculateAlignmentPositionAndTextOffsets(cell, x, y, width, height);
            Size = new Vector2(x + width - LocalPosition.X, y + height - LocalPosition.Y);
            CheckClipping(cell, width);
        }

        //private void HandleRichText(PdfPageSettings pageSettings, PdfDictionaries dictionaries, double maxWidth, double maxHeight)
        //{
        //    PdfWritingMode mode = cell.Style.TextRotation == 255 ? PdfWritingMode.VerticalTtb : PdfWritingMode.HorizontalLtr;
        //    List<TextToken> tokens = TokenizeRichText(cell.RichText, mode);

        //    var currentLine = new PdfCellTextLine();
        //    currentLine.WritingMode = mode;
        //    var lineAdvance = 0d;
        //    var lineCross = 0d;

        //    foreach (var token in tokens)
        //    {
        //        double tokenAdvance = token.Item.GlyphPositions.Sum(g => mode == PdfWritingMode.HorizontalLtr ? g.AdvanceX : g.AdvanceY);

        //        double tokenCross = token.Item.Ascent + token.Item.Descent;

        //        bool overflow = (mode == PdfWritingMode.HorizontalLtr && lineAdvance + tokenAdvance > maxWidth) ||
        //                        (mode == PdfWritingMode.VerticalTtb && lineAdvance + tokenAdvance > maxHeight);


        //        if (overflow)
        //        {
        //            if (token.IsWhitespace)
        //            {
        //                // commit current line
        //                CloseLine(currentLine, lineAdvance, lineCross);
        //                TextLines.Add(currentLine);

        //                // new line
        //                currentLine = NewLine(mode);
        //                lineAdvance = 0;
        //                lineCross = 0;
        //                continue;
        //            }

        //            // Case: token is a word → create new line
        //            CloseLine(currentLine, lineAdvance, lineCross);
        //            TextLines.Add(currentLine);

        //            currentLine = NewLine(mode);
        //            lineAdvance = 0;
        //            lineCross = 0;
        //        }

        //        // Add token to current line
        //        currentLine.TextItemCollection.Add(token.Item);
        //        lineAdvance += tokenAdvance;
        //        lineCross = Math.Max(lineCross, tokenCross);
        //    }
        //}

        //private PdfCellTextLine NewLine(PdfWritingMode mode)
        //{
        //    return new PdfCellTextLine { WritingMode = mode };
        //}

        //private void CloseLine(
        //    PdfCellTextLine line, double advance, double cross)
        //{
        //    line.Advance = advance;
        //    line.CrossSize = cross;
        //}

        //private void MeasureGlyphs(PdfCellTextItem item, PdfWritingMode mode)
        //{
        //    item.GlyphPositions = new List<GlyphPosition>();

        //    char prev = '\0';

        //    foreach (char c in item.Text)
        //    {
        //        font.FontFamily = item.FontName;
        //        font.Size = (float)item.FontSize;
        //        font.Style = ((cell.Style.Font.Bold ? MeasurementFontStyles.Bold : 0) |
        //                      (cell.Style.Font.Italic ? MeasurementFontStyles.Italic : 0) |
        //                      (cell.Style.Font.Strike ? MeasurementFontStyles.Strikeout : 0) |
        //                      (cell.Style.Font.UnderLine ? MeasurementFontStyles.Underline : 0))
        //                      switch
        //        {
        //            0 => MeasurementFontStyles.Regular,
        //            var s => s
        //        };
        //        var measurement = fontMeasurerTrueType.MeasureText(c.ToString(), font);

        //        double advance = measurement.Width;// GetGlyphAdvance(item.FontName, item.FontSize, c);
        //        double kerning = 0;// prev != '\0' ? GetKerning(item.FontName, item.FontSize, prev, c) : 0;

        //        double advX = mode == PdfWritingMode.HorizontalLtr ? measurement.Width + kerning : 0;
        //        double advY = mode == PdfWritingMode.VerticalTtb ? measurement.Height + kerning : 0;

        //        item.GlyphPositions.Add(new GlyphPosition
        //        {
        //            Character = c,
        //            AdvanceX = advX,
        //            AdvanceY = advY,
        //            OffsetX = 0,
        //            OffsetY = 0,
        //            GlyphBox = new Rect() { Width = 2,
        //                                    Height = 2,
        //            },
        //        });

        //        prev = c;
        //    }

        //    // Fill ascent/descent here
        //    //item.Ascent = GetAscent(item.FontName, item.FontSize);
        //    //item.Descent = GetDescent(item.FontName, item.FontSize);
        //}

        //private PdfCellTextItem CreateTextItem(ExcelRichText rt)
        //{
        //    return new PdfCellTextItem
        //    {
        //        FontName = rt.FontName,
        //        FontSize = rt.Size,
        //        Bold = rt.Bold,
        //        Italic = rt.Italic,
        //        Underline = rt.UnderLine,
        //        UnderlineType = rt.UnderLineType,
        //        Strike = rt.Strike,
        //        SuperScript = rt.VerticalAlign == ExcelVerticalAlignmentFont.Superscript,
        //        SubScript = rt.VerticalAlign == ExcelVerticalAlignmentFont.Subscript,
        //        FontColor = rt.Color,
        //        Text = rt.Text
        //    };
        //}

        //private List<TextToken> TokenizeRichText(ExcelRichTextCollection rich, PdfWritingMode mode)
        //{
        //    var result = new List<TextToken>();

        //    foreach (var rt in rich)
        //    {
        //        string text = rt.Text;
        //        int i = 0;

        //        while (i < text.Length)
        //        {
        //            // Whitespace run
        //            if (char.IsWhiteSpace(text[i]))
        //            {
        //                int start = i;
        //                while (i < text.Length && char.IsWhiteSpace(text[i]))
        //                    i++;

        //                string whitespace = text.Substring(start, i - start);

        //                var item = CreateTextItem(rt);
        //                item.Text = whitespace;
        //                MeasureGlyphs(item, mode);

        //                result.Add(new TextToken { IsWhitespace = true, Item = item });
        //            }
        //            // Word run
        //            else
        //            {
        //                int start = i;
        //                while (i < text.Length && !char.IsWhiteSpace(text[i]))
        //                    i++;

        //                string word = text.Substring(start, i - start);

        //                var item = CreateTextItem(rt);
        //                item.Text = word;
        //                MeasureGlyphs(item, mode);

        //                result.Add(new TextToken { IsWhitespace = false, Item = item });
        //            }
        //        }
        //    }

        //    return result;
        //}

        //class StyledRun
        //{
        //    public ExcelRichText SourceRich;  // original rich text source (for font, size, color, etc)
        //    public string Text = "";     // characters of this run
        //                                 // When measured, we will create a PdfCellTextItem for each run (or keep glyphs)
        //    public PdfCellTextItem MeasuredItem = null;
        //}

        //// Token is either whitespace token or a word token consisting of multiple StyledRuns
        //class TextToken
        //{
        //    public bool IsWhitespace;
        //    public List<StyledRun> Runs = new List<StyledRun>();
        //}

        //// Compare the style / formatting of two RichText items
        //bool SameStyle(ExcelRichText a, ExcelRichText b)
        //{
        //    if (a == null || b == null) return false;
        //    return a.FontName == b.FontName &&
        //           a.Size == b.Size &&
        //           a.Bold == b.Bold &&
        //           a.Italic == b.Italic &&
        //           a.UnderLine == b.UnderLine &&
        //           a.Strike == b.Strike &&
        //           a.VerticalAlign == b.VerticalAlign &&
        //           a.Color.Equals(b.Color) &&
        //           a.UnderLineType == b.UnderLineType;
        //}

        //// Tokenize RichTextCollection into word/whitespace tokens while preserving styling runs
        //List<TextToken> TokenizeRichTextPreserveRuns(ExcelRichTextCollection rich, PdfWritingMode mode)
        //{
        //    var tokens = new List<TextToken>();

        //    // current token state
        //    TextToken currentToken = null;
        //    StyledRun currentRun = null;

        //    foreach (var rt in rich)
        //    {
        //        string txt = rt.Text ?? "";
        //        for (int i = 0; i < txt.Length; i++)
        //        {
        //            char c = txt[i];
        //            bool isWsChar = char.IsWhiteSpace(c);

        //            // Start a new token if needed (first token or token type change)
        //            if (currentToken == null || currentToken.IsWhitespace != isWsChar)
        //            {
        //                currentToken = new TextToken { IsWhitespace = isWsChar };
        //                tokens.Add(currentToken);
        //                currentRun = null;
        //            }

        //            // If the last run exists and has same style, append; otherwise create new run
        //            if (currentRun == null || !SameStyle(currentRun.SourceRich, rt))
        //            {
        //                currentRun = new StyledRun { SourceRich = rt, Text = "" };
        //                currentToken.Runs.Add(currentRun);
        //            }

        //            currentRun.Text += c;
        //        }
        //    }

        //    // Remove any empty tokens (defensive)
        //    tokens.RemoveAll(t => t.Runs.Count == 0 || t.Runs.All(r => string.IsNullOrEmpty(r.Text)));

        //    // Measure runs into PdfCellTextItem.GlyphPositions so we can compute advances later.
        //    foreach (var token in tokens)
        //    {
        //        foreach (var run in token.Runs)
        //        {
        //            // Convert the run->PdfCellTextItem but leave glyph positions empty for now:
        //            var item = CreateTextItem(run.SourceRich); // you already have this helper
        //            item.Text = run.Text;

        //            // Measure glyphs for the run, store them in the run.MeasuredItem (so we only measure once)
        //            MeasureGlyphs(item, mode); // fills item.GlyphPositions, item.Ascent, item.Descent
        //            run.MeasuredItem = item;
        //        }
        //    }

        //    return tokens;
        //}
        //double TokenAdvance(TextToken token, PdfWritingMode mode)
        //{
        //    double adv = 0;
        //    foreach (var run in token.Runs)
        //    {
        //        var item = run.MeasuredItem;
        //        adv += item.GlyphPositions.Sum(g => mode == PdfWritingMode.HorizontalLtr ? g.AdvanceX : g.AdvanceY);
        //    }
        //    return adv;
        //}

        //public class GlyphPosIndex
        //{
        //    public GlyphPosition glyph;
        //    public int runIndex;
        //}

        //List<GlyphPosIndex> FlattenTokenGlyphs(TextToken token)
        //{
        //    var list = new List<GlyphPosIndex>();
        //    for (int r = 0; r < token.Runs.Count; r++)
        //    {
        //        var run = token.Runs[r];
        //        var item = run.MeasuredItem;
        //        foreach (var g in item.GlyphPositions)
        //            list.Add(new GlyphPosIndex() { glyph=g, runIndex=r } );
        //    }
        //    return list;
        //}
        //void SplitLongWordToken(TextToken token, List<PdfCellTextLine> result, ref PdfCellTextLine currentLine, ref double lineAdvance, ref double lineCross, double maxWidth, double maxHeight, PdfWritingMode mode)
        //{
        //    var flat = FlattenTokenGlyphs(token);
        //    int idx = 0;
        //    while (idx < flat.Count)
        //    {
        //        // create an accumulator for this broken-line chunk
        //        double adv = 0;
        //        double cross = 0;
        //        var chunkGlyphs = new List<GlyphPosIndex>();

        //        while (idx < flat.Count)
        //        {
        //            var candidate = flat[idx];
        //            double nextAdv = adv + (mode == PdfWritingMode.HorizontalLtr ? candidate.glyph.AdvanceX : candidate.glyph.AdvanceY);

        //            bool overflow = (mode == PdfWritingMode.HorizontalLtr && nextAdv > maxWidth) ||
        //                            (mode == PdfWritingMode.VerticalTtb && nextAdv > maxHeight);

        //            if (overflow && chunkGlyphs.Count > 0)
        //                break;

        //            // If overflow and chunkGlyphs is empty, we still allow at least one glyph (avoid infinite loop)
        //            chunkGlyphs.Add(candidate);
        //            adv = nextAdv;
        //            idx++;
        //        }

        //        // From chunkGlyphs, create PdfCellTextItems grouped by runIndex
        //        int pos = 0;
        //        while (pos < chunkGlyphs.Count)
        //        {
        //            int runIdx = chunkGlyphs[pos].runIndex;
        //            var run = token.Runs[runIdx];

        //            // gather consecutive glyphs with same runIdx
        //            var gpList = new List<GlyphPosition>();
        //            var textBuilder = new System.Text.StringBuilder();
        //            while (pos < chunkGlyphs.Count && chunkGlyphs[pos].runIndex == runIdx)
        //            {
        //                gpList.Add(chunkGlyphs[pos].glyph);
        //                textBuilder.Append(chunkGlyphs[pos].glyph.Character);
        //                pos++;
        //            }

        //            // create a PdfCellTextItem representing this piece
        //            var brokenItem = CloneWithoutText(run.MeasuredItem); // keep style and metrics
        //            brokenItem.Text = textBuilder.ToString();
        //            brokenItem.GlyphPositions = gpList;
        //            // optionally recompute ascent/descent for this brokenItem if needed
        //            currentLine.TextItemCollection.Add(brokenItem);
        //        }

        //        lineAdvance = adv;
        //        lineCross = token.Runs.Max(r => r.MeasuredItem.Ascent + r.MeasuredItem.Descent);

        //        // If more glyphs remain, close line
        //        if (idx < flat.Count)
        //        {
        //            CloseLine(currentLine, lineAdvance, lineCross);
        //            result.Add(currentLine);

        //            currentLine = NewLine(mode);
        //            lineAdvance = 0;
        //            lineCross = 0;
        //        }
        //    }
        //}








        //Handle rich text from cell.
        private void HandleRichText(PdfPageSettings pageSettings, PdfDictionaries dictionaries, double width, double height, double x, int rotation)
        {
            int i = 0;
            int j = 0;
            var text = string.Empty;
            PdfCellTextLine textLine = new PdfCellTextLine();
            textLine.IsRichText = true;
            if (rotation == 255)
            {
                double totalHeight = 0;
                while (i < cell.RichText.Count)
                {
                    var rt = cell.RichText[i];
                    font.FontFamily = rt.FontName;
                    font.Size = (float)rt.Size;
                    font.Style = ((cell.Style.Font.Bold ? MeasurementFontStyles.Bold : 0) |
                                  (cell.Style.Font.Italic ? MeasurementFontStyles.Italic : 0) |
                                  (cell.Style.Font.Strike ? MeasurementFontStyles.Strikeout : 0) |
                                  (cell.Style.Font.UnderLine ? MeasurementFontStyles.Underline : 0))
                                  switch
                    {
                        0 => MeasurementFontStyles.Regular,
                        var s => s
                    };
                    var measurement = fontMeasurerTrueType.MeasureText(rt.Text, font);
                    totalHeight += measurement.Height;
                    if (height < totalHeight && cell.Style.WrapText)
                    {
                        text = string.Empty;
                        while (j < i)
                        {
                            text += cell.RichText[j].Text;
                            j++;
                        }
                        textLine.Text = text;
                        j = i - 1;
                        TextLines.Add(textLine);
                        textLine = new PdfCellTextLine();
                        totalHeight = 0;
                    }
                    else
                    {
                        PdfCellTextItem textItem = new PdfCellTextItem();
                        textItem.Text = rt.Text;
                        textItem.FontName = rt.FontName;
                        textItem.FontFamily = rt.Family;
                        textItem.FontSize = rt.Size;
                        textItem.Bold = rt.Bold;
                        textItem.Italic = rt.Italic;
                        textItem.Strike = rt.Strike;
                        textItem.Underline = rt.UnderLine;
                        textItem.UnderlineType = rt.UnderLineType;
                        textItem.SuperScript = rt.VerticalAlign == ExcelVerticalAlignmentFont.Superscript;
                        textItem.SubScript = rt.VerticalAlign == ExcelVerticalAlignmentFont.Subscript;
                        textItem.FontColor = rt.Color;
                        textItem.TextLength = measurement.Width;
                        textItem.LineHeight = measurement.Height;
                        textItem.FontHeight = measurement.FontHeight;
                        var fontData = GetFontResourceData(dictionaries.Fonts, pageSettings, textItem);
                        double gbox = (fontData.Os2Table.sTypoAscender - fontData.Os2Table.sTypoDescender) * (cell.Style.Font.Size / fontData.HeadTable.UnitsPerEm);
                        textItem.GlyphBox.Width = gbox;
                        textItem.GlyphBox.Height = gbox;
                        textItem.SubFamily = "Regular";
                        if (textItem.Bold)
                        {
                            textItem.SubFamily = "Bold";
                            if (textItem.Italic)
                            {
                                textItem.SubFamily += " Italic";
                            }
                        }
                        else if (textItem.Italic)
                        {
                            textItem.SubFamily = "Italic";
                        }
                        GetFontResourceData(dictionaries.Fonts, pageSettings, textItem);
                        if (!textItem.characterOffset.ContainsKey(textItem.Text[0]))
                        {
                            var character = fontMeasurerTrueType.MeasureText(textItem.Text[0].ToString(), font);
                            var offset = x + (textItem.GlyphBox.Width - character.Width) / 2d;
                            offset = offset - x;
                            textItem.characterOffset.Add(textItem.Text[0], new Vector2(offset, 0));
                        }
                        textLine.TextItemCollection.Add(textItem);
                        i++;
                    }
                }
                TextLines.Add(textLine);
            }
            else
            {
                double totalWidth = 0;
                while (i < cell.RichText.Count)
                {
                    var rt = cell.RichText[i];
                    font.FontFamily = rt.FontName;
                    font.Size = (float)rt.Size;
                    font.Style = ((cell.Style.Font.Bold ? MeasurementFontStyles.Bold : 0) |
                                  (cell.Style.Font.Italic ? MeasurementFontStyles.Italic : 0) |
                                  (cell.Style.Font.Strike ? MeasurementFontStyles.Strikeout : 0) |
                                  (cell.Style.Font.UnderLine ? MeasurementFontStyles.Underline : 0))
                                  switch
                    {
                        0 => MeasurementFontStyles.Regular,
                        var s => s
                    };
                    var measurement = fontMeasurerTrueType.MeasureText(rt.Text, font);
                    totalWidth += measurement.Width;
                    if (width < totalWidth && cell.Style.WrapText)
                    {
                        text = string.Empty;
                        while (j < i)
                        {
                            text += cell.RichText[j].Text;
                            j++;
                        }
                        textLine.Text = text;
                        j = i - 1;
                        TextLines.Add(textLine);
                        textLine = new PdfCellTextLine();
                        totalWidth = 0;
                    }
                    else
                    {
                        PdfCellTextItem textItem = new PdfCellTextItem();
                        textItem.Text = rt.Text;
                        textItem.FontName = rt.FontName;
                        textItem.FontFamily = rt.Family;
                        textItem.FontSize = rt.Size;
                        textItem.Bold = rt.Bold;
                        textItem.Italic = rt.Italic;
                        textItem.Strike = rt.Strike;
                        textItem.Underline = rt.UnderLine;
                        textItem.UnderlineType = rt.UnderLineType;
                        textItem.SuperScript = rt.VerticalAlign == ExcelVerticalAlignmentFont.Superscript;
                        textItem.SubScript = rt.VerticalAlign == ExcelVerticalAlignmentFont.Subscript;
                        textItem.FontColor = rt.Color;
                        textItem.TextLength = measurement.Width;
                        textItem.LineHeight = measurement.Height;
                        textItem.FontHeight = measurement.FontHeight;
                        var fontData = GetFontResourceData(dictionaries.Fonts, pageSettings, textItem);
                        double gbox = (fontData.Os2Table.sTypoAscender - fontData.Os2Table.sTypoDescender) * (cell.Style.Font.Size / fontData.HeadTable.UnitsPerEm);
                        textItem.GlyphBox.Width = gbox;
                        textItem.GlyphBox.Height = gbox;
                        textItem.SubFamily = "Regular";
                        if (textItem.Bold)
                        {
                            textItem.SubFamily = "Bold";
                            if (textItem.Italic)
                            {
                                textItem.SubFamily += " Italic";
                            }
                        }
                        else if (textItem.Italic)
                        {
                            textItem.SubFamily = "Italic";
                        }
                        GetFontResourceData(dictionaries.Fonts, pageSettings, textItem);
                        textLine.TextItemCollection.Add(textItem);
                        i++;
                    }
                }
                text = string.Empty;
                while (j < i)
                {
                    text += cell.RichText[j].Text;
                    j++;
                }
                textLine.Text = text;
                TextLines.Add(textLine);
            }
        }

        //Handle text from cell.
        private void HandleText(PdfPageSettings pageSettings, PdfDictionaries dictionaries, double x, double y, double width, double height, int rotation)
        {
            var textItem = CreateTextItem();
            var fontData = GetFontResourceData(dictionaries.Fonts, pageSettings, textItem);
            font.FontFamily = textItem.FontName;
            font.Size = (float)textItem.FontSize;
            font.Style = ((cell.Style.Font.Bold ? MeasurementFontStyles.Bold : 0) |
                          (cell.Style.Font.Italic ? MeasurementFontStyles.Italic : 0) |
                          (cell.Style.Font.Strike ? MeasurementFontStyles.Strikeout : 0) |
                          (cell.Style.Font.UnderLine ? MeasurementFontStyles.Underline : 0))
                          switch
            {
                0 => MeasurementFontStyles.Regular,
                var s => s
            };
            var result = fontMeasurerTrueType.MeasureText(cell.Text, font);
            textItem.TextLength = result.Width;
            textItem.LineHeight = result.Height;
            textItem.FontHeight = result.FontHeight;
            double gbox = (fontData.Os2Table.sTypoAscender - fontData.Os2Table.sTypoDescender) * (cell.Style.Font.Size / fontData.HeadTable.UnitsPerEm);
            textItem.GlyphBox.Width = gbox;
            textItem.GlyphBox.Height = gbox;
            double TextHeight = 0d;
            PdfCellTextLine lineItem = new PdfCellTextLine();
            string lineText = string.Empty;
            int textLength = textItem.Text.Length;
            double lineHeight = textItem.LineHeight;
            if (cell.Style.WrapText)
            {
                if (rotation == 255)
                {
                    for (int i = 0; i < textLength; i++)
                    {
                        if (TextHeight + textItem.LineHeight >= height)
                        {
                            lineItem.Text = lineText.Trim();
                            textItem.Text = lineText.Trim();
                            result = fontMeasurerTrueType.MeasureText(lineText, font);
                            textItem.TextLength = result.Width;
                            textItem.LineHeight = result.Height;
                            textItem.FontHeight = result.FontHeight;
                            textItem.GlyphBox.Width = gbox;
                            textItem.GlyphBox.Height = gbox;
                            lineItem.TextItemCollection.Add(textItem);
                            TextLines.Add(lineItem);
                            textItem = CreateTextItem();
                            lineItem = new PdfCellTextLine();
                            TextHeight = 0;
                            lineText = string.Empty;
                        }
                        TextHeight += lineHeight;
                        lineText += textItem.Text[i];
                        if (!textItem.characterOffset.ContainsKey(textItem.Text[i]))
                        {
                            var character = fontMeasurerTrueType.MeasureText(textItem.Text[i].ToString(), font);
                            var offset = x + (textItem.GlyphBox.Width - character.Width) / 2d;
                            offset = offset - x;
                            textItem.characterOffset.Add(textItem.Text[i], new Vector2(offset, 0));
                        }
                    }
                    if (!string.IsNullOrEmpty(lineText))
                    {
                        lineItem.Text = lineText.Trim();
                        textItem.Text = lineText.Trim();
                        result = fontMeasurerTrueType.MeasureText(lineText, font);
                        textItem.TextLength = result.Width;
                        textItem.LineHeight = result.Height;
                        textItem.FontHeight = result.FontHeight;
                        textItem.GlyphBox.Width = gbox;
                        textItem.GlyphBox.Height = gbox;
                        lineItem.TextItemCollection.Add(textItem);
                        TextLines.Add(lineItem);
                    }
                }
                else
                {
                    var lines = fontMeasurerTrueType.MeasureAndWrapTextPoints(cell.Text, font, width);
                    foreach (var line in lines)
                    {
                        lineItem.Text = line.Trim();
                        textItem.Text = line.Trim();
                        result = fontMeasurerTrueType.MeasureText(line, font);
                        textItem.TextLength = result.Width;
                        textItem.LineHeight = result.Height;
                        textItem.FontHeight = result.FontHeight;
                        textItem.GlyphBox.Width = gbox;
                        textItem.GlyphBox.Height = gbox;
                        lineItem.TextItemCollection.Add(textItem);
                        TextLines.Add(lineItem);
                        textItem = CreateTextItem();
                        lineItem = new PdfCellTextLine();
                    }
                }
            }
            else
            {
                lineItem.Text = cell.Text;
                lineItem.TextItemCollection.Add(textItem);
                TextLines.Add(lineItem);
                for (int k = 0; k < textItem.Text.Length; k++)
                {
                    if (!textItem.characterOffset.ContainsKey(textItem.Text[k]))
                    {
                        var character = fontMeasurerTrueType.MeasureText(textItem.Text[k].ToString(), font);
                        var offset = x + (textItem.GlyphBox.Width - character.Width) / 2d;
                        offset = offset - x;
                        textItem.characterOffset.Add(textItem.Text[k], new Vector2(offset, 0));
                    }
                }
            }
        }

        private PdfCellTextItem CreateTextItem()
        {
            PdfCellTextItem textItem = new PdfCellTextItem();
            textItem.Text = cell.Text;
            textItem.FontName = cell.Style.Font.Name;
            textItem.FontFamily = cell.Style.Font.Family;
            textItem.FontSize = cell.Style.Font.Size;
            textItem.Bold = cell.Style.Font.Bold;
            textItem.Italic = cell.Style.Font.Italic;
            textItem.Strike = cell.Style.Font.Strike;
            textItem.Underline = cell.Style.Font.UnderLine;
            textItem.UnderlineType = cell.Style.Font.UnderLineType;
            textItem.SuperScript = cell.Style.Font.VerticalAlign == ExcelVerticalAlignmentFont.Superscript;
            textItem.SubScript = cell.Style.Font.VerticalAlign == ExcelVerticalAlignmentFont.Subscript;
            textItem.FontColor = PdfColor.SetColorFromHex(cell.Style.Font.Color.LookupColor());
            textItem.SubFamily = "Regular";
            if (textItem.Bold)
            {
                textItem.SubFamily = "Bold";
                if (textItem.Italic)
                {
                    textItem.SubFamily += " Italic";
                }
            }
            else if (textItem.Italic)
            {
                textItem.SubFamily = "Italic";
            }
            return textItem;
        }


        //Get font data from fontResources. If font does not exsist, add it to fontResources.
        private OpenTypeFont GetFontResourceData(Dictionary<string, PdfFontResource> fontResources, PdfPageSettings pageSettings, PdfCellTextItem FontData)
        {
            if (!fontResources.ContainsKey(FontData.FullFontName))
            {
                int label = 1;
                if (fontResources.Count > 0)
                {
                    label = fontResources.Last().Value.labelNumber + 1;
                }
                PdfFontResource fr = new PdfFontResource(FontData.FontName, FontData.SubFamily, label, pageSettings);
                fontResources.Add(FontData.FullFontName, fr);
                fontResources.Last().Value.fontData = PdfTextData.GetFontData(pageSettings, FontData.FontName, FontData.SubFamily);
            }
            return fontResources[FontData.FullFontName].fontData;
        }

        //Get the length of the longest line of text.
        private double GetLongestLine()
        {
            double length = 0;
            foreach (var line in TextLines)
            {
                length = line.TextLength > length ? line.TextLength : length;
            }
            return length;
        }

        //Calculate text position from alignment and offsets for each line of text.
        private Vector2 CalculateAlignmentPositionAndTextOffsets(ExcelRangeBase cell, double cellX, double CellY, double cellWidth, double cellHeight)
        {
            double x = 0d;
            double y = 0d;
            double textLength = GetLongestLine();
            double fontHeight = TextLines[0].FontHeight;
            double lineHeight = TextLines[0].LineHeight;
            switch (CellAlignmentData.HorizontalAlignment)
            {
                case ExcelHorizontalAlignment.Fill:
                case ExcelHorizontalAlignment.General:
                    if (double.TryParse(cell.Value.ToString(), out double value))
                    {
                        x = cellX + (cellWidth - textLength) - rightMargin;
                    }
                    else
                    {
                        x = cellX + rightMargin;
                    }
                    break;
                case ExcelHorizontalAlignment.Left:
                    x = cellX + rightMargin;
                    break;
                case ExcelHorizontalAlignment.Center:
                    x = cellX + (cellWidth - textLength) / 2d;
                    break;
                case ExcelHorizontalAlignment.Right:
                    x = cellX + (cellWidth - textLength) - rightMargin;
                    break;
            }
            switch (CellAlignmentData.VerticalAlignment)
            {
                case ExcelVerticalAlignment.Top:
                    y = (CellY + cellHeight) - (fontHeight / 2d) - bottomMargin;
                    break;
                case ExcelVerticalAlignment.Center:
                    // replaces Math.Clamp which didn't exist in the older frameworks.
                    //var min = CellY + bottomMargin;
                    //var max = CellY + cellHeight - bottomMargin;
                    //var val = CellY + cellHeight / 2d + fontHeight / 2d;
                    //if (val > max) { y = max; }
                    //else if (val < min) { y = min; }
                    //else { y = val; }
                    //y = System.Math.Clamp(CellY + cellHeight / 2d + FontData.Lines[0].FontHeight / 2d, CellY + bottomMargin, CellY + cellHeight - bottomMargin);
                    y = CellY + (cellHeight / 2d) - (lineHeight / 4d); ;
                    break;
                case ExcelVerticalAlignment.Bottom:
                    y = CellY + bottomMargin;
                    break;
            }
            if (CellAlignmentData.IsVertical)
            {
                //set textRotation to 0 and then set bool isVertical
                //In content stream check is vertical andr
                return new Vector2(x, y);
            }
            else if (CellAlignmentData.TextRotation < 0)
            {
                double rot = CellAlignmentData.TextRotation * System.Math.PI / 180.0;
                x += textLength * (1 - System.Math.Cos(rot));
                y -= textLength * System.Math.Sin(rot);
            }
            else if (CellAlignmentData.TextRotation > 0)
            {
                double rot = CellAlignmentData.TextRotation * System.Math.PI / 180.0;
                x += textLength * (1 - System.Math.Cos(rot));
            }

            var yOffset = 0d;
            for (int i = 1; i < TextLines.Count; i++)
            {
                yOffset += TextLines[i].LineHeight;
                switch (CellAlignmentData.HorizontalAlignment)
                {
                    case ExcelHorizontalAlignment.Fill:
                    case ExcelHorizontalAlignment.General:
                        if (double.TryParse(cell.Value.ToString(), out double value))
                        {
                            TextLines[i].Offset = -TextLines[i].TextLength;
                        }
                        else
                        {
                            TextLines[i].Offset = 0d;
                        }
                        break;
                    case ExcelHorizontalAlignment.Left:
                        TextLines[i].Offset = 0d;
                        break;
                    case ExcelHorizontalAlignment.Center:
                        TextLines[i].Offset = (cellX + (cellWidth - TextLines[i].TextLength) / 2d) - x;
                        break;
                    case ExcelHorizontalAlignment.Right:
                        TextLines[i].Offset = (cellX + (cellWidth - TextLines[i].TextLength) - rightMargin) - x;
                        break;
                }
            }
            return new Vector2(x, y + yOffset);
        }

        //Check if clipping is needed.
        private void CheckClipping(ExcelRangeBase cell, double width)
        {
            double textLength = 0d;
            foreach (var line in TextLines)
            {
                if (textLength < line.TextLength)
                {
                    textLength = line.TextLength;
                }
            }
            if (textLength >= width || cell.Merge)
            {
                if (cell.Merge ||
                   CellAlignmentData.HorizontalAlignment == ExcelHorizontalAlignment.Fill ||
                   CellAlignmentData.HorizontalAlignment == ExcelHorizontalAlignment.Left && cell.Worksheet.Cells[cell._fromRow, cell._fromCol + 1].Value != null ||
                   CellAlignmentData.HorizontalAlignment == ExcelHorizontalAlignment.Right && cell.Worksheet.Cells[cell._fromRow, cell._fromCol - 1 <= 0 ? 1 : cell._fromCol - 1].Value != null)
                {
                    Clip = true;
                }
            }
        }

        //Create clipping rectangle.
        internal void CreateClippingRect(List<Transform> cells)
        {
            if (Clip)
            {
                var cellName = Name.Split('_')[0];
                var pcc = cells.Where(x => x.Name.Contains(cellName)).Where(x => x is PdfCellLayout).ToList();
                if (pcc.Count > 0)
                {
                    Clipping = new Rect()
                    {
                        X = pcc[0].LocalPosition.X + rightMargin,
                        Y = pcc[0].LocalPosition.Y,
                        Width = pcc[0].Size.X - rightMargin * 2,
                        Height = pcc[0].Size.Y
                    };
                }
            }
        }
    }
}
