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
using EPPlus.Graphics.Colors;
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
    internal class PdfCellContentLayout : Transform, ILayout
    {
        public PdfCellFontData FontData;
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
            FontData = new PdfCellFontData();
            FontData.FontName = cell.Style.Font.Name;
            FontData.FontFamily = cell.Style.Font.Family;
            FontData.FontSize = cell.Style.Font.Size;
            FontData.Bold = cell.Style.Font.Bold;
            FontData.Italic = cell.Style.Font.Italic;
            FontData.Strike = cell.Style.Font.Strike;
            FontData.Underline = cell.Style.Font.UnderLine;
            FontData.UnderlineType = cell.Style.Font.UnderLineType;
            FontData.SuperScript = cell.Style.Font.VerticalAlign == ExcelVerticalAlignmentFont.Superscript;
            FontData.SubScript = cell.Style.Font.VerticalAlign == ExcelVerticalAlignmentFont.Subscript;
            FontData.FontColor = new Color(cell.Style.Font.Color.LookupColor());
            FontData.SubFamily = "Regular";
            if (FontData.Bold)
            {
                FontData.SubFamily = "Bold";
                if (FontData.Italic)
                {
                    FontData.SubFamily += " Italic";
                }
            }
            else if (FontData.Italic)
            {
                FontData.SubFamily = "Italic";
            }
            var otfont = GetFontResourceData(dictionaries.Fonts, pageSettings);
            font.FontFamily = FontData.FontName;
            font.Size = (float)FontData.FontSize;
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
            FontData.TextLength = result.Width;
            FontData.LineHeight = result.Height;
            FontData.FontHeight = result.FontHeight;
            if (width < FontData.TextLength && cell.Style.WrapText)
            {

                var lines = fontMeasurerTrueType.MeasureAndWrapTextPoints(cell.Text, font, width);
                foreach(var line in lines)
                {
                    PdfTextLine l = new PdfTextLine();
                    l.Text = line;
                    var r = fontMeasurerTrueType.MeasureText(line, font);
                    l.TextLength = r.Width;
                    l.LineHeight = r.Height;
                    l.FontHeight = r.FontHeight;
                    FontData.Lines.Add(l);
                }
            }
            else
            {
                PdfTextLine l = new PdfTextLine();
                l.Text = cell.Text;
                l.TextLength = result.Width;
                l.LineHeight = result.Height;
                l.FontHeight = result.FontHeight;
                FontData.Lines.Add(l);
            }
            CellAlignmentData = new PdfCellAlignmentData();
            CellAlignmentData.HorizontalAlignment = cell.Style.HorizontalAlignment;
            CellAlignmentData.VerticalAlignment = cell.Style.VerticalAlignment;
            CellAlignmentData.Indent = cell.Style.Indent;
            CellAlignmentData.WrapText = cell.Style.WrapText;
            CellAlignmentData.ShrinkToFit = cell.Style.ShrinkToFit; //Need to fix Transform issues and then implement a method that sets scale on the text object.
            CellAlignmentData.TextRotation = cell.Style.TextRotation; //EPPlus does probably not calculate cell width and height after setting rotation on text. So before we make pdf we need to calculate cell width and height based on text rotation
            CellAlignmentData.TextDirection = cell.Style.ReadingOrder;
            //LocalPosition = CalculatePosition(cell, x, y, width, height, FontData.TextLength, FontData.LineHeight);
            LocalPosition = CalculateAlignmentPositionAndTextOffsets(cell, x, y, width, height);
            Size = new Vector2(x + width - LocalPosition.X, y + height - LocalPosition.Y);
            CheckClipping(cell, FontData.TextLength, width);
        }

        //Get font data from fontResources. If font does not exsist, add it to fontResources.
        private OpenTypeFont GetFontResourceData(Dictionary<string, PdfFontResource> fontResources, PdfPageSettings pageSettings)
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

        private Vector2 CalculateAlignmentPositionAndTextOffsets(ExcelRangeBase cell, double cellX, double CellY, double cellWidth, double cellHeight)
        {
            double x = 0d;
            double y = 0d;
            switch (CellAlignmentData.HorizontalAlignment)
            {
                case ExcelHorizontalAlignment.Fill:
                case ExcelHorizontalAlignment.General:
                    if (double.TryParse(cell.Value.ToString(), out double value))
                    {
                        x = cellX + (cellWidth - FontData.Lines[0].TextLength) - rightMargin;
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
                    x = cellX + (cellWidth - FontData.Lines[0].TextLength) / 2d;
                    break;
                case ExcelHorizontalAlignment.Right:
                    x = cellX + (cellWidth - FontData.Lines[0].TextLength) - rightMargin;
                    break;
            }
            switch (CellAlignmentData.VerticalAlignment)
            {
                case ExcelVerticalAlignment.Top:
                    y = CellY + FontData.Lines[0].FontHeight - bottomMargin;
                    break;
                case ExcelVerticalAlignment.Center:
                    // replaces Math.Clamp which didn't exist in the older frameworks.
                    var min = CellY + bottomMargin;
                    var max = CellY + cellHeight - bottomMargin;
                    var val = CellY + cellHeight / 2d + FontData.Lines[0].FontHeight / 2d;
                    if (val > max) { y = max; }
                    else if (val < min) { y = min; }
                    else { y = val; }
                    //y = System.Math.Clamp(CellY + cellHeight / 2d + FontData.Lines[0].FontHeight / 2d, CellY + bottomMargin, CellY + cellHeight - bottomMargin);
                    break;
                case ExcelVerticalAlignment.Bottom:
                    y = CellY + (cellHeight - FontData.Lines[0].FontHeight) + FontData.Lines[0].FontHeight - bottomMargin;
                    break;
            }
            var yOffset = 0d;
            for (int i = 1; i < FontData.Lines.Count; i++)
            {
                yOffset += FontData.Lines[i].LineHeight;
                switch (CellAlignmentData.HorizontalAlignment)
                {
                    case ExcelHorizontalAlignment.Fill:
                    case ExcelHorizontalAlignment.General:
                        if (double.TryParse(cell.Value.ToString(), out double value))
                        {
                            FontData.Lines[i].Offset = -FontData.Lines[i].TextLength;
                        }
                        else
                        {
                            FontData.Lines[i].Offset = 0d;
                        }
                        break;
                    case ExcelHorizontalAlignment.Left:
                        FontData.Lines[i].Offset = 0d;
                        break;
                    case ExcelHorizontalAlignment.Center:
                        FontData.Lines[i].Offset = (cellX + (cellWidth - FontData.Lines[i].TextLength) / 2d) - x;
                        break;
                    case ExcelHorizontalAlignment.Right:
                        FontData.Lines[i].Offset = (cellX + (cellWidth - FontData.Lines[i].TextLength) - rightMargin) - x;
                        break;
                }
            }

            return new Vector2(x, y-yOffset);
        }


        //Calculate text position based on alignment and cell size.
        private Vector2 CalculatePosition(ExcelRangeBase cell, double cellX, double CellY, double cellWidth, double cellHeight, double textLength, double fontHeight)
        {
            /*take a list of strings
             * one line for each line
             * calculate text matrix
             * calculate offset for each line depending on alignment
             * 
             * Alignmet coords are calculated from cell left, right and top and bottom
             * Top for text needs to be at Y = top - firstRowLineHeight - margin
             * Right for text is X = left + width - margin = right - margin
             * Left is X = left + margin
             * Down is Y = down + margin
             * 
             * Next step is to take Line heights and text widths into consideration
             * For Center:
             * X = X + CellWidth / 2
             * X = X - TextWidth / 2
             * 
             * For Right:
             * X = X + cellWidth
             * X = X - TextWidth
             * 
             * For Left:
             * X = X
             * 
             * For Top:
             * Y = Y + CellHeight - FirstLineHeight
             * 
             * For Bottom:
             * Y = Y + AllLineHeightsExceptFirst
             * 
             * For Center:
             * diff = CellHeight - AllLineHeights
             * Y = Y + diff/2
             * 
             * WE need height of each line for Y placement
             * WE need length of each text line
             * We need to store the offset for PDF Td operator
             * 
             * KEEP IN MIND axis has not been flipped to bottom left yet. We are still in top right. Math may be adjusted
            */






            double x = 0d;
            double y = 0d;
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
                    y = CellY + fontHeight + bottomMargin;
                    break;
                case ExcelVerticalAlignment.Center:
                    y = CellY + (cellHeight - fontHeight) / 2d + fontHeight;
                    break;
                case ExcelVerticalAlignment.Bottom:
                    y = CellY + (cellHeight - fontHeight) + fontHeight - bottomMargin;
                    break;
            }
            return new Vector2(x, y);
        }

        //Check if clipping is needed.
        private void CheckClipping(ExcelRangeBase cell, double textLength, double width)
        {
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
        internal void CreateClippingRect(List<Transform> cells)
        {
            if (Clip)
            {
                var pcc = cells.Where(x => x.Name == Name).Where(x=> x is PdfCellLayout).ToList();
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

        public void ConvertCoordinates(PdfPageSettings pageSettings)
        {
            LocalPosition = new Vector2(LocalPosition.X, pageSettings.PageSize.HeightPu - System.Math.Abs(LocalPosition.Y));
            //Convert other coordinates
        }
    }
}
