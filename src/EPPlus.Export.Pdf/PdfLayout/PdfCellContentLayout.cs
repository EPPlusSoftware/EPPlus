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
using EPPlus.Export.Pdf.PdfCatalog;
using EPPlus.Export.Pdf.PdfResources;
using EPPlus.Export.Pdf.PdfSettings;
using EPPlus.Fonts.OpenType.Integration;
using EPPlus.Graphics;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Text;
using OfficeOpenXml.Interfaces.Fonts;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Xml.Serialization;
using Vector2 = EPPlus.Graphics.Geometry.Vector2;

namespace EPPlus.Export.Pdf.PdfLayout
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
        public PdfCellStyle CellStyle;
        public ExcelRangeBase cell;
        public TextLayoutEngine textLayoutEngine;
        //private List<PdfTextFormat> textFormats = new List<PdfTextFormat>();
        public TextLineCollection TextLines;
        public List<PdfShapedText> ShapedTexts { get; set; }

        private double textLength = 0;
        private double textHeight = 0;
        public double LeftTextSpillLength = 0d;
        public double RightTextSpillLength = 0d;
        private double bottomMargin = 3.5d; //Guessed number
        private double rightMargin = 1.4d; //I guessed this one too..
        //internal static FontMeasurerTrueType fontMeasurerTrueType = new FontMeasurerTrueType();
        //internal static MeasurementFont font = new MeasurementFont();

        //public List<PdfTextFormat> TextFormats { get => textFormats; set => textFormats = value; }
        public double TextLength { get => textLength; set => textLength = value; }
        public double TextHeight { get => textHeight; set => textHeight = value; }
        public TextLayoutEngine TextLayoutEngine { get => textLayoutEngine; set => textLayoutEngine = value; }



        public PdfCellContentLayout(PdfPageSettings pageSettings, PdfDictionaries dictionaries, PdfCell cell, MergedCellDrawInfo mergedCellInfo, double x, double y, double width, double height, double scaleX = 1, double scaleY = 1, double rotation = 0, Transform parent = null)
            : base(x, y-height, width, height, scaleX, scaleY, rotation, parent)
        {
            Z = 2;
            CellAlignmentData = cell.ContentAligmnet;
            TextLines = cell.TextLines;
            ShapedTexts = cell.ShapedTexts;
            TextLayoutEngine = cell.TextLayoutEngine;
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
            TextLayoutEngine = headerFooter.Content.TextLayoutEngine;
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
                    newX = x + padding;
                    break;
                case ExcelHorizontalAlignment.Center:
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
            double newX = CalculateHorizontalAlignment(text, textLength, x, width, rightMargin);
            double newY = CalculateVerticalAlignment(text, textHeight, firstLineAscent, lastLineAscent, y, height, 0d);
            return CalculatePositionFromRotation(textLength, newX, newY);
        }

        //private static List<TextFragment> GetTextFragments(List<PdfTextFormat> textFormats)
        //{
        //    var fragments = new List<TextFragment>(textFormats.Count);

        //    foreach (var tf in textFormats)
        //    {
        //        var fragment = new TextFragment
        //        {
        //            Text = tf.Text,
        //            Font = new MeasurementFont
        //            {
        //                FontFamily = tf.FontName,
        //                Size = (float)tf.FontSize,
        //                Style = GetMeasurementFontStyle(tf)
        //            }
        //        };
        //        fragments.Add(fragment);
        //    }

        //    return fragments;
        //}

        //private static MeasurementFontStyles GetMeasurementFontStyle(PdfTextFormat tf)
        //{
        //    var style = (tf.Bold ? MeasurementFontStyles.Bold : 0)
        //              | (tf.Italic ? MeasurementFontStyles.Italic : 0)
        //              | (tf.Strike ? MeasurementFontStyles.Strikeout : 0)
        //              | (tf.Underline ? MeasurementFontStyles.Underline : 0);

        //    return style == 0 ? MeasurementFontStyles.Regular : (MeasurementFontStyles)style;
        //}




        public PdfCellContentLayout(ExcelRangeBase cell, PdfCellStyle CellStyle, PdfPageSettings pageSettings, double x, double y, double width, double height, double scaleX = 1, double scaleY = 1, double rotation = 0, Transform parent = null, PdfDictionaries dictionaries = null)
            : base(x, y, width, height, scaleX, scaleY, rotation, parent)
        {
            this.cell = cell;
            this.CellStyle = CellStyle;
            CellAlignmentData = new PdfCellAlignmentData();
            CellAlignmentData.HorizontalAlignment = cell.Style.HorizontalAlignment;
            CellAlignmentData.VerticalAlignment = cell.Style.VerticalAlignment;
            CellAlignmentData.Indent = cell.Style.Indent;
            CellAlignmentData.WrapText = cell.Style.WrapText;
            CellAlignmentData.ShrinkToFit = cell.Style.ShrinkToFit;
            CellAlignmentData.TextRotation = (cell.Style.TextRotation >= 90) ? ((cell.Style.TextRotation == 255) ? 0 : 90 - cell.Style.TextRotation) : cell.Style.TextRotation;
            CellAlignmentData.IsVertical = cell.Style.TextRotation == 255 ? true : false;
            CellAlignmentData.TextDirection = cell.Style.ReadingOrder;
            if (!cell.IsRichText) cell._rtc = new ExcelRichTextCollection(cell.Text, cell);
            //HandleText(pageSettings, dictionaries, x, y, width, height, CellAlignmentData.TextRotation, CellStyle);
            //CalculateTextSpill(width, CellAlignmentData.TextRotation);
            //LocalPosition = CalculateAlignmentPositionAndTextOffsets(cell, x, y, width, height);
            Size = new Vector2(x + width - LocalPosition.X, y + height - LocalPosition.Y); 
            //CheckClipping(cell, x, y, width, height);
        }

        public PdfCellContentLayout(string text, ExcelRangeBase cell, PdfCellStyle CellStyle, PdfPageSettings pageSettings, double x, double y, double width, double height, double scaleX = 1, double scaleY = 1, double rotation = 0, Transform parent = null, PdfDictionaries dictionaries = null)
            : base(x, y, width, height, scaleX, scaleY, rotation, parent)
        {
            this.cell = cell;
            this.CellStyle = CellStyle;
            CellAlignmentData = new PdfCellAlignmentData();
            CellAlignmentData.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            CellAlignmentData.VerticalAlignment = ExcelVerticalAlignment.Bottom;
            cell._rtc = new ExcelRichTextCollection(text, cell);
            var ns = cell.Worksheet.Workbook.Styles.GetNormalStyle();
            cell._rtc[0].FontName = ns.Style.Font.Name;
            cell._rtc[0].Family = ns.Style.Font.Family;
            cell._rtc[0].Size = ns.Style.Font.Size;
            cell._rtc[0].Color = Color.Black;
            //HandleText(pageSettings, dictionaries, x, y, width, height, CellAlignmentData.TextRotation, CellStyle);
            //CalculateTextSpill(width, CellAlignmentData.TextRotation);
            //LocalPosition = CalculateAlignmentPositionAndTextOffsets(cell, x, y, width, height);
            Size = new Vector2(x + width - LocalPosition.X, y + height - LocalPosition.Y);
            //CheckClipping(cell, x, y, width, height);
        }

        //private void HandleText(PdfPageSettings pageSettings, PdfDictionaries dictionaries, double x, double y, double maxWidth, double maxHeight, double rotation, PdfCellStyle CellStyle)
        //{
        //    bool bold = false, italic = false, underline = false, strike = false;
        //    ExcelUnderLineType underLineType = ExcelUnderLineType.None;
        //    if (CellStyle != null && CellStyle.dxfFont != null)
        //    {
        //        bold = CellStyle.dxfFont.Bold != null ? (bool)CellStyle.dxfFont.Bold : false;
        //        italic = CellStyle.dxfFont.Italic != null ? (bool)CellStyle.dxfFont.Italic : false;
        //        strike = CellStyle.dxfFont.Strike != null ? (bool)CellStyle.dxfFont.Strike : false;
        //        underline = CellStyle.dxfFont.Underline != null;
        //        underLineType = CellStyle.dxfFont.Underline != null ? (ExcelUnderLineType)CellStyle.dxfFont.Underline : ExcelUnderLineType.None;
        //    }
        //    for (int i = 0; i < cell.RichText.Count; i++)
        //    {
        //        var rt = cell.RichText[i];
        //        var textformat = new PdfTextFormat();
        //        textformat.Text = rt.Text;
        //        textformat.FontName = rt.FontName;
        //        textformat.FontFamily = rt.Family;
        //        textformat.FontSize = rt.Size;
        //        textformat.Bold = rt.Bold || bold;
        //        textformat.Italic = rt.Italic || italic;
        //        textformat.Strike = rt.Strike || strike;
        //        textformat.Underline = rt.UnderLine || underline;
        //        textformat.UnderlineType = rt.UnderLineType == ExcelUnderLineType.None ? underLineType : rt.UnderLineType;
        //        textformat.SuperScript = rt.VerticalAlign == ExcelVerticalAlignmentFont.Superscript;
        //        textformat.SubScript = rt.VerticalAlign == ExcelVerticalAlignmentFont.Subscript;
        //        textformat.FontColor = rt.Color;
        //        textformat.SubFamily = FontSubFamily.Regular;
        //        if (textformat.Bold)
        //        {
        //            textformat.SubFamily = FontSubFamily.Bold;
        //            if (textformat.Italic)
        //            {
        //                textformat.SubFamily = FontSubFamily.BoldItalic;
        //            }
        //        }
        //        else if (textformat.Italic)
        //        {
        //            textformat.SubFamily = FontSubFamily.Italic;
        //        }
               
        //        this.textFormats.Add(textformat);
        //        if (!dictionaries.Fonts.ContainsKey(textformat.FullFontName))
        //        {
        //            int label = 1;
        //            if (dictionaries.Fonts.Count > 0)
        //            {
        //                label = dictionaries.Fonts.Last().Value.labelNumber + 1;
        //            }
        //            dictionaries.Fonts.Add(textformat.FullFontName, new PdfFontResource(textformat.FontName, textformat.SubFamily, label, pageSettings));
        //        }
        //        var manger = dictionaries.Fonts[textformat.FullFontName].fontSubsetManager;
        //        manger.AddText(textformat.Text);
        //    }
        //}

        public void CalculateTextSpill(double maxWidth, double rotation)
        {
            if (maxWidth < textLength)
            {
                if (rotation != 0)
                {
                    var baseVec = new Vector2(textLength, 0);
                    var rad = rotation * Math.PI / 180.0d;
                    var rotVec = new Vector2(textLength * Math.Cos(rad), textLength * Math.Sin(rad));
                    var length = Vector2.Project(rotVec, baseVec).Length;
                    if (length > maxWidth)
                    {
                        if (CellAlignmentData.HorizontalAlignment == ExcelHorizontalAlignment.General)
                        {
                            if (double.TryParse(cell.Value.ToString(), out double value))
                            {
                                LeftTextSpillLength = textLength;
                            }
                            else
                            {
                                RightTextSpillLength = textLength;
                            }
                        }
                        else if (CellAlignmentData.HorizontalAlignment == ExcelHorizontalAlignment.Left)
                        {
                            RightTextSpillLength = length;
                        }
                        else if (CellAlignmentData.HorizontalAlignment == ExcelHorizontalAlignment.Right)
                        {
                            LeftTextSpillLength = length;
                        }
                        else if (CellAlignmentData.HorizontalAlignment == ExcelHorizontalAlignment.Center)
                        {
                            LeftTextSpillLength = length / 2d;
                            RightTextSpillLength = length / 2d;
                        }

                    }
                    return;
                }
                if (CellAlignmentData.HorizontalAlignment == ExcelHorizontalAlignment.General)
                {
                    if (double.TryParse(cell.Value.ToString(), out double value))
                    {
                        LeftTextSpillLength = textLength - Size.X;
                    }
                    else
                    {
                        RightTextSpillLength = textLength - Size.X;
                    }
                }
                else if (CellAlignmentData.HorizontalAlignment == ExcelHorizontalAlignment.Left)
                {
                    RightTextSpillLength = textLength - Size.X;
                }
                else if (CellAlignmentData.HorizontalAlignment == ExcelHorizontalAlignment.Right)
                {
                    LeftTextSpillLength = textLength - Size.X;
                }
                else if (CellAlignmentData.HorizontalAlignment == ExcelHorizontalAlignment.Center)
                {
                    LeftTextSpillLength = (textLength - Size.X) / 2d;
                    RightTextSpillLength = (textLength - Size.X) / 2d;
                }
            }
        }

        //Calculate text position from alignment and offsets for each line of text.
        public Vector2 CalculateAlignmentPositionAndTextOffsets(ExcelRangeBase cell, double cellX, double CellY, double cellWidth, double cellHeight)
        {
            double x = 0d;
            double y = 0d;
            double xOffset = 0d;
            double yOffset = 0d;
            //double textLength = Lines.TextLength;
            //double textHeight = Lines.TextHeight;
            //double fontHeight = Lines.FontHeight;
            //double lineHeight = Lines.LineHeight;
            //if (CellAlignmentData.IsVertical)
            //{
                //Need to place vertical text a bit better. It appears each character is monospaced when using vertical. so we should define a common glyphbox that we then use for placing and measuring vertical text.
                //switch (CellAlignmentData.HorizontalAlignment)
                //{
                    //case ExcelHorizontalAlignment.Left:
                    //    x = cellX + rightMargin;
                    //    break;
                    //case ExcelHorizontalAlignment.Fill:
                    //case ExcelHorizontalAlignment.General:
                    //case ExcelHorizontalAlignment.Center:
                    //    x = cellX + (cellWidth - Lines.Height) / 2d;
                    //    break;
                    //case ExcelHorizontalAlignment.Right:
                    //    x = cellX + (cellWidth - Lines.Height) - rightMargin;
                    //    break;
                //}
                //switch (CellAlignmentData.VerticalAlignment)
                //{
                //    case ExcelVerticalAlignment.Top:
                //        y = (CellY + cellHeight) - (fontHeight / 2d) - bottomMargin;
                //        break;
                //    case ExcelVerticalAlignment.Center:
                //        y = CellY + (cellHeight / 2d) + (Lines.Lines[0].TextHeight / 4d);
                //        break;
                //    case ExcelVerticalAlignment.Bottom:
                //        y = CellY + (Lines.Lines[0].TextHeight - Lines.Lines[0].Words[0].Characters[0].LineHeight) + bottomMargin;
                //        break;
                //}
                //for (int i = 1; i < Lines.Lines.Count; i++)
                //{
                //    //xOffset += Lines.LineHeight;
                //    switch (CellAlignmentData.VerticalAlignment)
                //    {
                //        case ExcelVerticalAlignment.Top:
                //            Lines.Lines[i].Offset = 0d;
                //            break;
                //        case ExcelVerticalAlignment.Center:
                //            Lines.Lines[i].Offset = CellY + (cellHeight / 2d) + (Lines.Lines[i].TextHeight / 4d) - y;
                //            break;
                //        case ExcelVerticalAlignment.Bottom:
                //            Lines.Lines[i].Offset = (CellY + (Lines.Lines[i].TextHeight - Lines.Lines[i].Words[0].Characters[0].LineHeight) + bottomMargin) - y;
                //            break;
                //    }
                //}
            //}
            //else
            //{
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
                        y = (CellY + cellHeight) - (textHeight / 2d) - bottomMargin;
                        break;
                    case ExcelVerticalAlignment.Center:
                        y = CellY + (cellHeight / 2d) - (textHeight / 4d); ;
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
                //for (int i = 1; i < Lines.Lines.Count; i++)
                //{
                //    yOffset += Lines.LineHeight;
                //    switch (CellAlignmentData.HorizontalAlignment)
                //    {
                //        case ExcelHorizontalAlignment.Fill:
                //        case ExcelHorizontalAlignment.General:
                //            if (double.TryParse(cell.Value.ToString(), out double value))
                //            {
                //                Lines.Lines[i].Offset = -Lines.Lines[i].TextLength;
                //            }
                //            else
                //            {
                //                Lines.Lines[i].Offset = 0d;
                //            }
                //            break;
                //        case ExcelHorizontalAlignment.Left:
                //            Lines.Lines[i].Offset = 0d;
                //            break;
                //        case ExcelHorizontalAlignment.Center:
                //            Lines.Lines[i].Offset = (cellX + (cellWidth - Lines.Lines[i].TextLength) / 2d) - x;
                //            break;
                //        case ExcelHorizontalAlignment.Right:
                //            Lines.Lines[i].Offset = (cellX + (cellWidth - Lines.Lines[i].TextLength) - rightMargin) - x;
                //            break;
                //    }
                //}
            //}
            return new Vector2(x, y);
        }

        //Check if clipping is needed.


        //Check if clipping is needed.

        public void CheckClipping(ExcelRangeBase cell, double x, double y, double width, double height)
        {
            if (textLength >= width || cell.Merge)
            {
                if (cell.Merge ||
                   CellAlignmentData.HorizontalAlignment == ExcelHorizontalAlignment.Fill)
                //CellAlignmentData.HorizontalAlignment == ExcelHorizontalAlignment.Left && cell.Worksheet.Cells[cell._fromRow, cell._fromCol + 1].Value != null ||
                //CellAlignmentData.HorizontalAlignment == ExcelHorizontalAlignment.Right && cell.Worksheet.Cells[cell._fromRow, cell._fromCol - 1 <= 0 ? 1 : cell._fromCol - 1].Value != null)
                {
                    Clipping = new Rect()
                    {
                        X = x + rightMargin,
                        Y = y,
                        Width = width - rightMargin * 2,
                        Height = height
                    };
                }
            }
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

        //Create clipping rectangle.
        internal void CreateClippingRect(PdfCellLayout cell, double targetX)
        {
            //need a way to detect starting X position if text is spilled from left or right
            var width = targetX - cell.LocalPosition.X;
            Clipping = new Rect()
            {
                X = cell.LocalPosition.X + rightMargin,
                Y = cell.LocalPosition.Y,
                Width =  width - rightMargin * 2,
                Height = cell.Size.Y
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

                    dictionaries.Fonts[font.FullName].Gids.Add(glyph.GlyphId);
                    dictionaries.Fonts[font.FullName].fontData = font;

                    if (!dictionaries.Fonts[font.FullName].charactermappings.ContainsKey(glyph.GlyphId))
                    {
                        var chars = ExtractCharactersForGlyph(glyph, tf.ShapedText.OriginalText);
                        if (!string.IsNullOrEmpty(chars))
                        {
                            dictionaries.Fonts[font.FullName].charactermappings[glyph.GlyphId] = chars;
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
