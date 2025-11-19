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
using EPPlus.Export.Pdf.Math;
using EPPlus.Export.Pdf.PdfGraphics;
using EPPlus.Export.Pdf.Pdfhelpers;
using EPPlus.Export.Pdf.PdfResources;
using EPPlus.Export.Pdf.PdfSettings;
using EPPlus.Fonts.OpenType;
using OfficeOpenXml;
using OfficeOpenXml.Interfaces.Drawing.Text;
using OfficeOpenXml.Style;
using System.Collections.Generic;
using System.Linq;

namespace EPPlus.Export.Pdf.PdfLayout
{
    internal class PdfCellContentLayout : PdfTransform, ILayout
    {
        public List<PdfCellTextLine> TextLines = new List<PdfCellTextLine>();
        public PdfCellAlignmentData CellAlignmentData;
        public bool Clip;
        public PdfRect Clipping;

        private double bottomMargin = 3.5d; //Guessed number
        private double rightMargin = 1.4d; //I guessed this one too..

        internal static FontMeasurerTrueType fontMeasurerTrueType = new FontMeasurerTrueType();
        internal static MeasurementFont font = new MeasurementFont();

        public PdfCellContentLayout(ExcelRangeBase cell, PdfPageSettings pageSettings, double x, double y, double width, double height, double scaleX = 1, double scaleY = 1, double rotation = 0, PdfTransform parent = null, PdfDictionaries dictionaries = null)
            : base(x, y, width, height, scaleX, scaleY, rotation, parent)
        {
            this.cell = cell;
            if (cell.IsRichText)
            {
                HandleRichText(pageSettings, dictionaries, width, cell.Style.TextRotation);
            }
            else
            {
                HandleText(pageSettings, dictionaries, width, height, cell.Style.TextRotation);
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

        //Handle rich text from cell.
        private void HandleRichText(PdfPageSettings pageSettings, PdfDictionaries dictionaries, double width, int rotation)
        {
            int i = 0;
            int j = 0;
            double totalWidth = 0;
            var text = string.Empty;
            PdfCellTextLine textLine = new PdfCellTextLine();
            textLine.IsRichText = true;
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
                    textItem.FontColor = new PdfColor(rt.Color.R, rt.Color.G, rt.Color.B, rt.Color.A);
                    textItem.TextLength = measurement.Width;
                    textItem.LineHeight = measurement.Height;
                    textItem.FontHeight = measurement.FontHeight;
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

        /*
         * Measure text in width and height
         * if rotation is 255
         * measure the line of text for each cahracter and sum lineheight
         * do the same thing we do now for width but height instead
         */


        //Handle text from cell.
        private void HandleText(PdfPageSettings pageSettings, PdfDictionaries dictionaries, double width, double height, int rotation)
        {
            var textItem = CreateTextItem();
            GetFontResourceData(dictionaries.Fonts, pageSettings, textItem);
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
                            lineItem.TextItemCollection.Add(textItem);
                            TextLines.Add(lineItem);
                            textItem = CreateTextItem();
                            lineItem = new PdfCellTextLine();
                            TextHeight = 0;
                            lineText = string.Empty;
                        }
                        TextHeight += lineHeight;
                        lineText += textItem.Text[i];
                    }
                    if (!string.IsNullOrEmpty(lineText))
                    {
                        lineItem.Text = lineText.Trim();
                        textItem.Text = lineText.Trim();
                        result = fontMeasurerTrueType.MeasureText(lineText, font);
                        textItem.TextLength = result.Width;
                        textItem.LineHeight = result.Height;
                        textItem.FontHeight = result.FontHeight;
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
            textItem.FontColor = new PdfColor(cell.Style.Font.Color.LookupColor());
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
                    y = CellY + fontHeight - bottomMargin;
                    break;
                case ExcelVerticalAlignment.Center:
                    // replaces Math.Clamp which didn't exist in the older frameworks.
                    var min = CellY + bottomMargin;
                    var max = CellY + cellHeight - bottomMargin;
                    var val = CellY + cellHeight / 2d + fontHeight / 2d;
                    if (val > max) { y = max; }
                    else if (val < min) { y = min; }
                    else { y = val; }
                    //y = System.Math.Clamp(CellY + cellHeight / 2d + FontData.Lines[0].FontHeight / 2d, CellY + bottomMargin, CellY + cellHeight - bottomMargin);
                    break;
                case ExcelVerticalAlignment.Bottom:
                    y = CellY + (cellHeight - fontHeight) + fontHeight - bottomMargin;
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
                y += textLength * System.Math.Sin(rot);
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
            return new Vector2(x, y-yOffset);
        }

        //Check if clipping is needed.
        private void CheckClipping(ExcelRangeBase cell, double width)
        {
            double textLength = 0d;
            foreach (var line in TextLines)
            {
                if(textLength < line.TextLength)
                {
                    textLength = line.TextLength;
                }
            }
            if (textLength >= width || cell.Merge)
            {
                if ( cell.Merge ||
                   CellAlignmentData.HorizontalAlignment == ExcelHorizontalAlignment.Fill ||
                   CellAlignmentData.HorizontalAlignment == ExcelHorizontalAlignment.Left  && cell.Worksheet.Cells[cell._fromRow, cell._fromCol + 1].Value != null ||
                   CellAlignmentData.HorizontalAlignment == ExcelHorizontalAlignment.Right && cell.Worksheet.Cells[cell._fromRow, cell._fromCol - 1 <= 0 ? 1 : cell._fromCol - 1].Value != null)
                {
                    Clip = true;
                }
            }
        }

        //Create clipping rectangle.
        internal void CreateClippingRect(List<PdfTransform> cells)
        {
            if (Clip)
            {
                var pcc = cells.Where(x => x.Name == Name).Where(x=> x is PdfCellLayout).ToList();
                if (pcc.Count > 0)
                {
                    Clipping = new PdfRect()
                    {
                        X = pcc[0].LocalPosition.X + rightMargin,
                        Y = pcc[0].LocalPosition.Y,
                        Width = pcc[0].Size.X - rightMargin * 2,
                        Height = pcc[0].Size.Y
                    };
                }
            }
        }

        //Convert to pdf coordinates.
        public void ConvertCoordinates(PdfPageSettings pageSettings)
        {
            LocalPosition = new Vector2(LocalPosition.X, pageSettings.PageSize.HeightPu - System.Math.Abs(LocalPosition.Y));
        }
    }
}
