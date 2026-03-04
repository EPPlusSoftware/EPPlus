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
using EPPlus.Export.Pdf.Pdfhelpers;
using EPPlus.Export.Pdf.PdfLayout;
using EPPlus.Export.Pdf.PdfResources;
using EPPlus.Export.Pdf.PdfSettings;
using EPPlus.Fonts.OpenType;
using EPPlus.Graphics;
using EPPlus.Graphics.Math;
using OfficeOpenXml;
using OfficeOpenXml.Interfaces.Fonts;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;

namespace EPPlus.Export.Pdf.PdfObjects
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
            else if (cell.CellFillData.BackgroundColor != Color.Empty && cell.CellFillData.PatternStyle != ExcelFillStyle.None)
            {
                commands.Add($"% Gradient Start: {cell.Name}");
                commands.Add("q");
                commands.Add($"{GridLine.HalfWidth.ToPdfString()} w");
                commands.Add(cell.CellFillData.BackgroundColor.ToFillCommand());
                commands.Add(cell.CellFillData.BackgroundColor.ToStrokeCommand());
                commands.Add($"{cell.LocalPosition.X.ToPdfString()} {cell.LocalPosition.Y.ToPdfString()} {cell.Size.X.ToPdfString()} {cell.Size.Y.ToPdfString()} re");
                commands.Add("B");
                commands.Add("Q");
                commands.Add("q");
                commands.Add("/Pattern cs");
                commands.Add($"/{label} scn");
                commands.Add($"{cell.LocalPosition.X.ToPdfString()} {cell.LocalPosition.Y.ToPdfString()} {cell.Size.X.ToPdfString()} {cell.Size.Y.ToPdfString()} re");
                commands.Add("f");
                commands.Add("Q");
                commands.Add($"% Gradient End: {cell.Name}");
            }
        }

        public void AddBorderLayout(PdfCellBorderLayout cell)
        {
            var borderRenderer = new PdfBorderRenderer(cell);
            borderRenderer.RenderBorder(this);
        }


        //Get font label //need to update this one too for same reasons as AddFontData
        internal PdfFontResource GetFontResource(PdfDictionaries Dictionaries, PdfPageSettings PageSettings, string fontName, FontSubFamily subFamily, double fontSize)
        {
            if (!Dictionaries.Fonts.ContainsKey(fontName))
            {
                int label = 1;
                if (Dictionaries.Fonts.Count > 0)
                {
                    label = Dictionaries.Fonts.Last().Value.labelNumber + 1;
                }
                PdfFontResource fr = new PdfFontResource(fontName, subFamily, label, PageSettings);
                if (fontName != "Courier New")
                {
                    //Document.Add(fr.GetFontDescriptorObject(Document.Count + 1));
                    //Document.Add(fr.GetWidthsObject(Document.Count + 1));
                }
                //Document.Add(fr.GetFontObject(Document.Count + 1));
                Dictionaries.Fonts.Add(fontName, fr);
            }
            return Dictionaries.Fonts[fontName];
        }

        public void AddText(PdfCellContentLayout cell, PdfDictionaries dictionaries, PdfPageSettings pageSettings)
        {
            var position = cell.LocalPosition;
            var alignment = cell.CellAlignmentData;
            double advanceX = 0;

            double rotation = alignment.TextRotation * System.Math.PI / 180.0;
            //bool isVertical = alignment.IsVertical;

            for (int i = 0; i < cell.fontData.Count; i++)
            {
                byte currentFontId = 0;



                var fontData = cell.fontData[i];
                var text = cell.ShapedText[i];
                var textLength = text.GetWidthInPoints((float)fontData.FontSize, 2048);
                var color = fontData.FontColor;
                var font = GetFontResource(dictionaries, pageSettings, fontData.FullFontName, fontData.SubFamily, fontData.FontSize);
                double size = fontData.FontSize;
                double scale = fontData.FontSize / font.fontData.HeadTable.UnitsPerEm;
                Matrix3x3 textMatrix = new Matrix3x3(System.Math.Cos(rotation), System.Math.Sin(rotation), -System.Math.Sin(rotation), System.Math.Cos(rotation), position.X, position.Y);
                commands.Add("BT");
                textMatrix = textMatrix * Matrix3x3.Translation(advanceX, 0);
                if (fontData.SuperScript)
                {
                    var supOffX = font.fontData.Os2Table.ySuperscriptXOffset * scale;
                    var supOffY = font.fontData.Os2Table.ySuperscriptYOffset * scale;
                    var supSizeY = font.fontData.Os2Table.ySuperscriptYSize * scale;
                    textMatrix = textMatrix * Matrix3x3.Translation(supOffX, supOffY);
                    size = supSizeY;
                }
                else if (fontData.SubScript)
                {
                    var supOffX = font.fontData.Os2Table.ySubscriptXOffset * scale;
                    var supOffY = font.fontData.Os2Table.ySubscriptYOffset * scale;
                    var supSizeY = font.fontData.Os2Table.ySubscriptYSize * scale;
                    textMatrix = textMatrix * Matrix3x3.Translation(supOffX, supOffY);
                    size = supSizeY;
                }
                if (fontData.Underline)
                {
                    var underlinePos = font.fontData.PostTable.underlinePosition * scale;
                    var underlineWidth = font.fontData.PostTable.underlineThickness * scale;
                    var start = textMatrix.Transform(new Vector2(0, underlinePos));
                    var end = textMatrix.Transform(new Vector2(textLength, underlinePos));
                    commands.Add($"{underlineWidth.ToPdfString()} w");
                    commands.Add($"{start.X.ToPdfString()} {start.Y.ToPdfString()} m");
                    commands.Add($"{end.X.ToPdfString()} {end.Y.ToPdfString()} l");
                    commands.Add($"S");
                }
                if (fontData.Strike)
                {
                    var strikePos = font.fontData.Os2Table.yStrikeoutPosition * scale;
                    var strikeWidth = font.fontData.Os2Table.yStrikeoutSize * scale;
                    var start = textMatrix.Transform(new Vector2(0, strikePos));
                    var end = textMatrix.Transform(new Vector2(textLength, strikePos));
                    commands.Add($"{strikeWidth.ToPdfString()} w");
                    commands.Add($"{start.X.ToPdfString()} {start.Y.ToPdfString()} m");
                    commands.Add($"{end.X.ToPdfString()} {end.Y.ToPdfString()} l");
                    commands.Add($"S");
                }
                commands.Add(color.ToFillCommand());
                commands.Add($"{textMatrix.A.ToPdfString()} {textMatrix.B.ToPdfString()} {textMatrix.C.ToPdfString()} {textMatrix.D.ToPdfString()} {textMatrix.E.ToPdfString()} {textMatrix.F.ToPdfString()} Tm");
                commands.Add($"/{font.Label} {size.ToPdfString()} Tf"); //move to inside for looop
                var sb = new StringBuilder();
                sb.Append("[");
                for (int j = 0; j < text.Glyphs.Length; j++)
                {
                    var glyph = text.Glyphs[j];

                    if (glyph.FontId != currentFontId)
                    {
                        sb.Append("] TJ");
                        sb.Append($"/{font.Label} {size.ToPdfString()} Tf");
                        sb.Append("[");
                    }

                    sb.Append($"<{glyph.GlyphId:X4}>");
                    int kerning = glyph.XAdvance - glyph.BaseAdvance;

                    if (kerning != 0)
                    {
                        // Convert to PDF units (1000-based) and negate
                        // PDF uses negative values to ADD space, positive to REMOVE space
                        double adjustment = -(kerning * 1000.0 / 1000);
                        sb.Append($" {adjustment.ToPdfStringF0()}");
                    }

                    if (j < text.Glyphs.Length - 1)
                    {
                        sb.Append(" ");
                    }
                }
                advanceX += textLength;
                commands.Add(sb.ToString() + "] TJ");
                commands.Add("ET");
            }
        }

        public void AddText(Vector2 position, PdfCellLines lines, PdfCellAlignmentData alignment, PdfDictionaries dictionaries, PdfPageSettings pageSettings)
        {
            double rot = alignment.TextRotation * System.Math.PI / 180.0;
            bool isVertical = alignment.IsVertical;
            Matrix3x3 textMatrix = new Matrix3x3(System.Math.Cos(rot), System.Math.Sin(rot), -System.Math.Sin(rot), System.Math.Cos(rot), position.X, position.Y);
            for (int i = 0; i < lines.Lines.Count; i++)
            {
                var line = lines.Lines[i];
                Matrix3x3 textRunMatrix = textMatrix;
                Matrix3x3 modifierMatrix = Matrix3x3.Identity;
                bool useModifiedMatrix = false;
                commands.Add("BT");
                commands.Add($"{textMatrix.A.ToPdfString()} {textMatrix.B.ToPdfString()} {textMatrix.C.ToPdfString()} {textMatrix.D.ToPdfString()} {textMatrix.E.ToPdfString()} {textMatrix.F.ToPdfString()} Tm");
                PdfCellTextItem lastCharacter = line.Words[0].Characters[0];
                PdfCellTextItem currentStyle = line.Words[0].Characters[0];
                string textRun = string.Empty;
                double textAdvance = 0d;
                double textVAdvance = 0d;
                int wordIndex = 0;
                for (int j = 0; j < line.Words.Count; j++)
                {
                    var words = line.Words[j];

                    for (int k = wordIndex; k < words.Characters.Count; k++, wordIndex++)
                    {
                        if (!currentStyle.Equals(words.Characters[k]))
                        {
                            break;
                        }
                        textRun += words.Characters[k].Text;
                        textAdvance += words.Characters[k].TextLength;
                        textVAdvance = words.Characters[k].LineHeight;
                        if (isVertical)
                        {
                            wordIndex++;
                            break;
                        }
                    }

                    if (wordIndex == words.Characters.Count && j < line.Words.Count - 1)
                    {
                        wordIndex = 0;
                        continue;
                    }

                    var font = GetFontResource(dictionaries, pageSettings, currentStyle.FullFontName, currentStyle.SubFamily, currentStyle.FontSize);
                    double size = currentStyle.FontSize;
                    double scale = currentStyle.FontSize / font.fontData.HeadTable.UnitsPerEm;
                    if (currentStyle.Bold)
                    {
                        commands.Add("0.25 w");
                        commands.Add("2 Tr");
                        commands.Add(currentStyle.FontColor.ToStrokeCommand());
                    }
                    else
                    {
                        commands.Add("0 Tr");
                    }
                    if (currentStyle.Italic)
                    {
                        var ia = font.fontData.PostTable.italicAngle.FloatValue;
                        if (ia <= 0) ia = 12f * (float)System.Math.PI / 180.0f;
                        modifierMatrix.C = System.Math.Tan(ia);
                        modifierMatrix = modifierMatrix * textRunMatrix;
                        useModifiedMatrix = true;
                    }
                    if (currentStyle.SuperScript)
                    {
                        var supOffX = font.fontData.Os2Table.ySuperscriptXOffset * scale;
                        var supOffY = font.fontData.Os2Table.ySuperscriptYOffset * scale;
                        var supSizeX = font.fontData.Os2Table.ySuperscriptXSize * scale;
                        var supSizeY = font.fontData.Os2Table.ySuperscriptYSize * scale;
                        modifierMatrix.E = textMatrix.E + supOffX;
                        modifierMatrix.F = textMatrix.F + supOffY;
                        size = supSizeY;
                        useModifiedMatrix = true;
                    }
                    else if (currentStyle.SubScript)
                    {
                        var supOffX = font.fontData.Os2Table.ySubscriptXOffset * scale;
                        var supOffY = font.fontData.Os2Table.ySubscriptYOffset * scale;
                        var supSizeX = font.fontData.Os2Table.ySubscriptXSize * scale;
                        var supSizeY = font.fontData.Os2Table.ySubscriptYSize * scale;
                        modifierMatrix.E = textMatrix.E + supOffX;
                        modifierMatrix.F = textMatrix.F + supOffY;
                        size = supSizeY;
                        useModifiedMatrix = true;
                    }
                    if (currentStyle.Underline)
                    {
                        var underlinePos = font.fontData.PostTable.underlinePosition * scale;
                        var underlineWidth = font.fontData.PostTable.underlineThickness * scale;
                        var start = textRunMatrix.Transform(new Vector2(0, underlinePos));
                        var end = textRunMatrix.Transform(new Vector2(textAdvance, underlinePos));
                        commands.Add($"{underlineWidth.ToPdfString()} w");
                        commands.Add($"{start.X.ToPdfString()} {start.Y.ToPdfString()} m");
                        commands.Add($"{end.X.ToPdfString()} {end.Y.ToPdfString()} l");
                        commands.Add($"S");
                    }
                    if (currentStyle.Strike)
                    {
                        var strikePos = font.fontData.Os2Table.yStrikeoutPosition * scale;
                        var strikeWidth = font.fontData.Os2Table.yStrikeoutSize * scale;
                        var start = textRunMatrix.Transform(new Vector2(0, strikePos));
                        var end = textRunMatrix.Transform(new Vector2(textAdvance, strikePos));
                        commands.Add($"{strikeWidth.ToPdfString()} w");
                        commands.Add($"{start.X.ToPdfString()} {start.Y.ToPdfString()} m");
                        commands.Add($"{end.X.ToPdfString()} {end.Y.ToPdfString()} l");
                        commands.Add($"S");
                    }
                    if (useModifiedMatrix)
                    {
                        commands.Add($"{modifierMatrix.A.ToPdfString()} {modifierMatrix.B.ToPdfString()} {modifierMatrix.C.ToPdfString()} {modifierMatrix.D.ToPdfString()} {modifierMatrix.E.ToPdfString()} {modifierMatrix.F.ToPdfString()} Tm");
                        modifierMatrix = Matrix3x3.Identity;
                    }
                    else if ((isVertical))
                    {
                        commands.Add($"{textRunMatrix.A.ToPdfString()} {textRunMatrix.B.ToPdfString()} {textRunMatrix.C.ToPdfString()} {textRunMatrix.D.ToPdfString()} {textRunMatrix.E.ToPdfString()} {textRunMatrix.F.ToPdfString()} Tm");
                    }
                    commands.Add($"/{font.Label} {size.ToPdfString()} Tf");
                    commands.Add(currentStyle.FontColor.ToFillCommand());
                    commands.Add($"({FixEscapeCharacters(textRun)}) Tj");
                    if (isVertical)
                    {
                        textRunMatrix = textRunMatrix * Matrix3x3.Translation(0, -textVAdvance);
                    }
                    else
                    {
                        textRunMatrix = textRunMatrix * Matrix3x3.Translation(textAdvance, 0);
                    }

                    if (useModifiedMatrix) commands.Add($"{textRunMatrix.A.ToPdfString()} {textRunMatrix.B.ToPdfString()} {textRunMatrix.C.ToPdfString()} {textRunMatrix.D.ToPdfString()} {textRunMatrix.E.ToPdfString()} {textRunMatrix.F.ToPdfString()} Tm");
                    useModifiedMatrix = false;
                    textRun = string.Empty;
                    textAdvance = 0;

                    if (wordIndex < words.Characters.Count)
                    {
                        currentStyle = words.Characters[wordIndex];
                        j--;
                    }
                }
                if (i + 1 < lines.Lines.Count)
                {
                    if (isVertical)
                    {
                        textMatrix = textRunMatrix * Matrix3x3.Translation(line.TextLength, line.TextHeight + lines.Lines[i + 1].Offset);
                    }
                    else
                    {
                        textMatrix = textRunMatrix * Matrix3x3.Translation(-line.TextLength + lines.Lines[i + 1].Offset, -lines.Lines[i + 1].LineHeight);
                    }
                }
                commands.Add("ET");
            }
        }

        public void AddCellContentLayout(PdfCellContentLayout cell, PdfDictionaries dictionaries, PdfPageSettings pageSettings)
        {
            commands.Add($"% Content Start: {cell.Name}");
            commands.Add("q");
            if (cell.Clip) AddClipping(cell);
            AddText(cell, dictionaries, pageSettings);
            commands.Add("Q");
            commands.Add($"% Content End: {cell.Name}");
        }

        private void AddClipping(PdfCellContentLayout cell)
        {
            commands.Add($"{cell.Clipping.X.ToPdfString()} {cell.Clipping.Y.ToPdfString()} {cell.Clipping.Width.ToPdfString()} {cell.Clipping.Height.ToPdfString()} re W n");
        }

        public void AddCellContentLayout(PdfHeaderFooterLayout cell, PdfDictionaries dictionaries, PdfPageSettings pageSettings)
        {
            commands.Add($"% HeaderFooter Start: {cell.Name}");
            commands.Add("q");
            //AddText(cell.LocalPosition, cell.Lines, cell.CellAlignmentData, dictionaries, pageSettings);
            commands.Add("Q");
            commands.Add($"% HeaderFooter End: {cell.Name}");
        }

        public void AddInnerGridLines(Transform pageLayout)
        {
            if (pageLayout is not PdfPageLayout pl) return;

            commands.Add($"% Gridlnes Start");
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

        public void AddOuterGridBorder(Transform pageLayout)
        {
            if (pageLayout is not PdfPageLayout pl) return;

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

        public void AddMarginClipping(Transform pageLayout, PdfContentBounds bounds)
        {
            if (pageLayout is not PdfPageLayout pl) return;

            commands.Add($"% Margin Clip Start");
            double y = bounds.Top;
            double width = 0d;
            foreach (var line in pl.BorderLines)
            {
                width = System.Math.Max(width, System.Math.Max(line.X1, line.X2));
                y = System.Math.Min(y, System.Math.Min(line.Y1, line.Y2));
            }
            var heightAdjust = y - bounds.Bottom;
            commands.Add($"{bounds.X.ToPdfString()} {y.ToPdfString()} {(width - bounds.Left).ToPdfString()} {(bounds.Height - heightAdjust).ToPdfString()} re W n");
            //commands.Add($"{bounds.X.ToPdfString()} {bounds.Y.ToPdfString()} {bounds.Width.ToPdfString()} {bounds.Height.ToPdfString()} re W n");
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
            var bytes = Encoding.ASCII.GetBytes(content);
            WriteAscii(bw, $"<< /Length {bytes.Length} >>\nstream\n{content}\nendstream");
        }

        private string FixEscapeCharacters(string text)
        {
            return text.Replace(@"\", @"\\").Replace(@"(", @"\(").Replace(@")", @"\)");
        }
    }
}
