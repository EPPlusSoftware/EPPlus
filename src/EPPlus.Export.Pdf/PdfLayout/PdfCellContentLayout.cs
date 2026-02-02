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
using EPPlus.Export.Pdf.Pdfhelpers;
using EPPlus.Export.Pdf.PdfResources;
using EPPlus.Export.Pdf.PdfSettings;
using EPPlus.Fonts.OpenType;
using EPPlus.Graphics;
using EPPlus.Graphics.Math;
using OfficeOpenXml;
using OfficeOpenXml.Interfaces.Drawing.Text;
using OfficeOpenXml.Style;
using System.Collections.Generic;
using System.Linq;

namespace EPPlus.Export.Pdf.PdfLayout
{
    internal class PdfCellContentLayout : Transform
    {
        public PdfCellLines Lines = new PdfCellLines();
        public PdfCellAlignmentData CellAlignmentData;
        public bool Clip;
        public Rect Clipping;
        public PdfCellStyleOverride CellStyle;

        private double bottomMargin = 3.5d; //Guessed number
        private double rightMargin = 1.4d; //I guessed this one too..

        internal static FontMeasurerTrueType fontMeasurerTrueType = new FontMeasurerTrueType();
        internal static MeasurementFont font = new MeasurementFont();

        public PdfCellContentLayout(ExcelRangeBase cell, PdfCellStyleOverride CellStyle, PdfPageSettings pageSettings, double x, double y, double width, double height, double scaleX = 1, double scaleY = 1, double rotation = 0, Transform parent = null, PdfDictionaries dictionaries = null)
            : base(x, y, width, height, scaleX, scaleY, rotation, parent)
        {
            this.cell = cell;
            this.CellStyle = CellStyle;
            if (cell.IsRichText)
            {
                //HandleRichText(pageSettings, dictionaries, width, height, x, cell.Style.TextRotation);
                HandleRichText(pageSettings, dictionaries, x, y, width, height, cell.Style.TextRotation, CellStyle);
            }
            else
            {
                cell._rtc = new ExcelRichTextCollection(cell.Text, cell);
                HandleRichText(pageSettings, dictionaries, x, y, width, height, cell.Style.TextRotation, CellStyle);
                //HandleText(pageSettings, dictionaries, x, y, width, height, cell.Style.TextRotation);
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
            LocalPosition = CalculateAlignmentPositionAndTextOffsets(cell, x, y, width, height);
            Size = new Vector2(x + width - LocalPosition.X, y + height - LocalPosition.Y);
            CheckClipping(cell, width);
        }

        private void HandleRichText(PdfPageSettings pageSettings, PdfDictionaries dictionaries, double x, double y, double maxWidth, double maxHeight, double rotation, PdfCellStyleOverride CellStyle)
        {
            List<PdfCellTextItem> characters = new List<PdfCellTextItem>();
            List<PdfCellWord> Words = new List<PdfCellWord>();
            bool bold = false, italic = false, underline = false, strike = false;
            ExcelUnderLineType underLineType = ExcelUnderLineType.None;
            if(CellStyle.dxfFont != null)
            {
                bold = CellStyle.dxfFont.Bold != null ? (bool)CellStyle.dxfFont.Bold : false;
                italic = CellStyle.dxfFont.Italic != null ? (bool)CellStyle.dxfFont.Italic : false;
                strike = CellStyle.dxfFont.Strike != null ? (bool)CellStyle.dxfFont.Strike : false;
                underline = CellStyle.dxfFont.Underline != null;
                underLineType = CellStyle.dxfFont.Underline != null ? (ExcelUnderLineType)CellStyle.dxfFont.Underline : ExcelUnderLineType.None;
            }
            for (int i = 0; i < cell.RichText.Count; i++)
            {
                var rt = cell.RichText[i];
                for (int j = 0; j < rt.Text.Length; j++)
                {
                    var character = new PdfCellTextItem();
                    character.Text = rt.Text[j].ToString();
                    character.FontName = rt.FontName;
                    character.FontFamily = rt.Family;
                    character.FontSize = rt.Size;
                    character.Bold = rt.Bold || bold;
                    character.Italic = rt.Italic || italic;
                    character.Strike = rt.Strike || strike;
                    character.Underline = rt.UnderLine || underline;
                    character.UnderlineType = rt.UnderLineType == ExcelUnderLineType.None ? underLineType : rt.UnderLineType;
                    character.SuperScript = rt.VerticalAlign == ExcelVerticalAlignmentFont.Superscript;
                    character.SubScript = rt.VerticalAlign == ExcelVerticalAlignmentFont.Subscript;
                    character.FontColor = rt.Color;
                    character.SubFamily = FontSubFamily.Regular;
                    if (character.Bold)
                    {
                        character.SubFamily = FontSubFamily.Bold;
                        if (character.Italic)
                        {
                            character.SubFamily = FontSubFamily.BoldItalic;
                        }
                    }
                    else if (character.Italic)
                    {
                        character.SubFamily = FontSubFamily.Italic;
                    }
                    font.FontFamily = character.FontName;
                    font.Size = (float)character.FontSize;
                    font.Style = ((cell.Style.Font.Bold ? MeasurementFontStyles.Bold : 0) |
                                  (cell.Style.Font.Italic ? MeasurementFontStyles.Italic : 0) |
                                  (cell.Style.Font.Strike ? MeasurementFontStyles.Strikeout : 0) |
                                  (cell.Style.Font.UnderLine ? MeasurementFontStyles.Underline : 0))
                                  switch
                    {
                        0 => MeasurementFontStyles.Regular,
                        var s => s
                    };
                    var result = fontMeasurerTrueType.MeasureText(character.Text, font);
                    character.TextLength = result.Width;
                    character.LineHeight = result.Height;
                    character.FontHeight = result.FontHeight;
                    var fontData = GetFontResourceData(dictionaries.Fonts, pageSettings, character);
                    double gbox = (fontData.Os2Table.sTypoAscender - fontData.Os2Table.sTypoDescender) * (cell.Style.Font.Size / fontData.HeadTable.UnitsPerEm);
                    character.GlyphBox = new Rect();
                    character.GlyphBox.Width = gbox;
                    character.GlyphBox.Height = gbox;
                    character.characterOffset = (x + (character.GlyphBox.Width - result.Width) / 2d) - x;
                    characters.Add(character);
                }
            }
            var w = new PdfCellWord();
            for (int i = 0; i < characters.Count; i++)
            {
                var t = characters[i];
                if (char.IsWhiteSpace(t.Text[0]))
                {
                    Words.Add(w);
                    w = new PdfCellWord();
                    w.Characters.Add(t);
                    Words.Add(w);
                    w = new PdfCellWord();
                }
                else
                {
                    w.Characters.Add(t);
                }
                if (i == characters.Count - 1)
                {
                    Words.Add(w);
                }
            }
            if (rotation == 255)
            {
                Lines.WritingMode = PdfWritingMode.VerticalTtb;
                if (cell.Style.WrapText)
                {
                    double lineHeight = 0d;
                    var line = new PdfCellLine();
                    for (int i = 0; i < Words.Count; i++)
                    {
                        if (Words[i].TextHeight > maxHeight)
                        {
                            double lineCount = Words[i].TextHeight / maxHeight;
                            int currentIndex = 0;
                            while (lineCount > 0)
                            {
                                PdfCellWord newWord = new PdfCellWord();
                                double wordlength = 0;
                                for (int j = currentIndex; j < Words[i].Characters.Count; j++)
                                {
                                    if (wordlength + Words[i].Characters[j].LineHeight > maxHeight)
                                    {
                                        currentIndex = j;
                                        break;
                                    }
                                    else
                                    {
                                        newWord.Characters.Add(Words[i].Characters[j]);
                                        wordlength += Words[i].Characters[j].LineHeight;
                                    }
                                }
                                line.Words.Add(newWord);
                                Lines.Lines.Add(line);
                                lineCount--;
                                line = new PdfCellLine();
                            }
                        }
                        else if (lineHeight + Words[i].TextHeight < maxHeight)
                        {
                            lineHeight += Words[i].TextHeight;
                            line.Words.Add(Words[i]);
                        }
                        else
                        {
                            Lines.Lines.Add(line);
                            lineHeight = Words[i].TextHeight;
                            line = new PdfCellLine();
                            line.Words.Add(Words[i]);
                        }
                    }
                    Lines.Lines.Add(line);
                }
                else
                {
                    var line = new PdfCellLine();
                    foreach (var word in Words)
                    {
                        line.Words.Add(word);
                    }
                    Lines.Lines.Add(line);
                }
            }
            else
            {
                Lines.WritingMode = PdfWritingMode.HorizontalLtr;
                if (cell.Style.WrapText)
                {
                    double lineLength = 0d;
                    var line = new PdfCellLine();
                    for (int i = 0; i < Words.Count; i++)
                    {
                        if (Words[i].TextLength > maxWidth)
                        {
                            double lineCount = Words[i].TextLength / maxWidth;
                            int currentIndex = 0;
                            while (lineCount > 0)
                            {
                                PdfCellWord newWord = new PdfCellWord();
                                double wordlength = 0;
                                for (int j = currentIndex; j < Words[i].Characters.Count; j++)
                                {
                                    if (wordlength + Words[i].Characters[j].TextLength > maxWidth)
                                    {
                                        currentIndex = j;
                                        break;
                                    }
                                    else
                                    {
                                        newWord.Characters.Add(Words[i].Characters[j]);
                                        wordlength += Words[i].Characters[j].TextLength;
                                    }
                                }
                                line.Words.Add(newWord);
                                Lines.Lines.Add(line);
                                lineCount--;
                                line = new PdfCellLine();
                            }
                        }
                        else if (lineLength + Words[i].TextLength < maxWidth)
                        {
                            lineLength += Words[i].TextLength;
                            line.Words.Add(Words[i]);
                        }
                        else
                        {
                            Lines.Lines.Add(line);
                            lineLength = Words[i].TextLength;
                            line = new PdfCellLine();
                            line.Words.Add(Words[i]);
                        }
                    }
                    Lines.Lines.Add(line);
                }
                else
                {
                    var line = new PdfCellLine();
                    foreach (var word in Words)
                    {
                        line.Words.Add(word);
                    }
                    Lines.Lines.Add(line);
                }
            }
            Lines.Lines.RemoveAll(l => l.Words.Count <= 0);
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

        //Calculate text position from alignment and offsets for each line of text.
        private Vector2 CalculateAlignmentPositionAndTextOffsets(ExcelRangeBase cell, double cellX, double CellY, double cellWidth, double cellHeight)
        {
            double x = 0d;
            double y = 0d;
            double xOffset = 0d;
            double yOffset = 0d;
            double textLength = Lines.TextLength;
            double textHeight = Lines.TextHeight;
            double fontHeight = Lines.FontHeight;
            double lineHeight = Lines.LineHeight;
            if (CellAlignmentData.IsVertical)
            {
                //Need to place vertical text a bit better. It appears each character is monospaced when using vertical. so we should define a common glyphbox that we then use for placing and measuring vertical text.
                switch (CellAlignmentData.HorizontalAlignment)
                {
                    case ExcelHorizontalAlignment.Left:
                        x = cellX + rightMargin;
                        break;
                    case ExcelHorizontalAlignment.Fill:
                    case ExcelHorizontalAlignment.General:
                    case ExcelHorizontalAlignment.Center:
                        x = cellX + (cellWidth - Lines.Height) / 2d;
                        break;
                    case ExcelHorizontalAlignment.Right:
                        x = cellX + (cellWidth - Lines.Height) - rightMargin;
                        break;
                }
                switch (CellAlignmentData.VerticalAlignment)
                {
                    case ExcelVerticalAlignment.Top:
                        y = (CellY + cellHeight) - (fontHeight / 2d) - bottomMargin;
                        break;
                    case ExcelVerticalAlignment.Center:
                        y = CellY + (cellHeight / 2d) + (Lines.Lines[0].TextHeight / 4d);
                        break;
                    case ExcelVerticalAlignment.Bottom:
                        y = CellY + (Lines.Lines[0].TextHeight - Lines.Lines[0].Words[0].Characters[0].LineHeight) + bottomMargin;
                        break;
                }
                for (int i = 1; i < Lines.Lines.Count; i++)
                {
                    //xOffset += Lines.LineHeight;
                    switch (CellAlignmentData.VerticalAlignment)
                    {
                        case ExcelVerticalAlignment.Top:
                            Lines.Lines[i].Offset = 0d;
                            break;
                        case ExcelVerticalAlignment.Center:
                            Lines.Lines[i].Offset = CellY + (cellHeight / 2d) + (Lines.Lines[i].TextHeight / 4d) - y;
                            break;
                        case ExcelVerticalAlignment.Bottom:
                            Lines.Lines[i].Offset = (CellY + (Lines.Lines[i].TextHeight - Lines.Lines[i].Words[0].Characters[0].LineHeight) + bottomMargin) - y;
                            break;
                    }
                }
            }
            else
            {
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
                        y = CellY + (cellHeight / 2d) - (lineHeight / 4d); ;
                        break;
                    case ExcelVerticalAlignment.Bottom:
                        y = CellY + bottomMargin;
                        break;
                }
                if (CellAlignmentData.TextRotation < 0)
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
                for (int i = 1; i < Lines.Lines.Count; i++)
                {
                    yOffset += Lines.LineHeight;
                    switch (CellAlignmentData.HorizontalAlignment)
                    {
                        case ExcelHorizontalAlignment.Fill:
                        case ExcelHorizontalAlignment.General:
                            if (double.TryParse(cell.Value.ToString(), out double value))
                            {
                                Lines.Lines[i].Offset = -Lines.Lines[i].TextLength;
                            }
                            else
                            {
                                Lines.Lines[i].Offset = 0d;
                            }
                            break;
                        case ExcelHorizontalAlignment.Left:
                            Lines.Lines[i].Offset = 0d;
                            break;
                        case ExcelHorizontalAlignment.Center:
                            Lines.Lines[i].Offset = (cellX + (cellWidth - Lines.Lines[i].TextLength) / 2d) - x;
                            break;
                        case ExcelHorizontalAlignment.Right:
                            Lines.Lines[i].Offset = (cellX + (cellWidth - Lines.Lines[i].TextLength) - rightMargin) - x;
                            break;
                    }
                }
            }
            return new Vector2(x + xOffset, y + yOffset);
        }

        //Check if clipping is needed.
        private void CheckClipping(ExcelRangeBase cell, double width)
        {
            if (Lines.TextLength >= width || cell.Merge)
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
