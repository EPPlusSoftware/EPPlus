using FontLab1;
using OfficeOpenXml.PDF.Math;
using OfficeOpenXml.PDF.PdfFontData;
using OfficeOpenXml.PDF.PdfGraphics;
using OfficeOpenXml.PDF.Pdfhelpers;
using OfficeOpenXml.PDF.PdfSettings;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography.Xml;

namespace OfficeOpenXml.PDF.PdfLayout
{

    internal class PdfCellContentLayout : PdfTransform
    {
        public PdfCellFontData FontData;
        public PdfCellAlignmentData CellAlignmentData;
        public readonly bool Clip;
        public PdfRect Clipping;

        private double bottomMargin = 3.4d; //Guessed number
        private double rightMargin = 1.0d; //I guessed this one too..

        public PdfCellContentLayout(ExcelRangeBase cell, PdfPageSettings pageSettings, double x, double y, double width, double height, double scaleX = 1, double scaleY = 1, double rotation = 0, PdfTransform parent = null, Dictionary<string, PdfFontResource> fontResources = null)
            : base(x, y, width, height, scaleX, scaleY, rotation, parent)
        {
            FontData = new PdfCellFontData();
            FontData.FontName = cell.Style.Font.Name;
            FontData.FontFamily = cell.Style.Font.Family;
            FontData.FontSize = cell.Style.Font.Size;
            FontData.Bold = cell.Style.Font.Bold;
            FontData.Italic = cell.Style.Font.Italic;
            FontData.Strike = cell.Style.Font.Strike;
            FontData.Underline = cell.Style.Font.UnderLine;
            FontData.UnderlineType = cell.Style.Font.UnderLineType;
            FontData.FontColor = new PdfColor(cell.Style.Font.Color.LookupColor());
            FontData.Text = cell.Text;
            CellAlignmentData = new PdfCellAlignmentData();
            CellAlignmentData.HorizontalAlignment = cell.Style.HorizontalAlignment;
            CellAlignmentData.VerticalAlignment = cell.Style.VerticalAlignment;
            CellAlignmentData.Indent = cell.Style.Indent;
            CellAlignmentData.WrapText = cell.Style.WrapText;
            CellAlignmentData.ShrinkToFit = cell.Style.ShrinkToFit; //Need to fix Transform issues and then implement a method that sets scale on the text object.
            CellAlignmentData.TextRotation = cell.Style.TextRotation; //EPPlus does probably not calculate cell width and height after setting rotation on text. So before we make pdf we need to calculate cell width and height based on text rotation
            CellAlignmentData.TextDirection = cell.Style.ReadingOrder;

            var ttfont = CacheFont(FontData, fontResources, pageSettings);
            double textLength = PdfTextData.MeasureText(FontData.Text, FontData.FontSize, ttfont);
            double fontHeight = PdfTextData.MeasureFontHeight(ttfont, FontData.FontSize);
            LocalPosition = CalculatePosition(cell, x, y, width, height, textLength, fontHeight);
            Size = new Vector2((x + width) - LocalPosition.X, (y + height) - LocalPosition.Y);
            if (textLength >= width)
            {
                if (CellAlignmentData.HorizontalAlignment == Style.ExcelHorizontalAlignment.Fill ||
                   (CellAlignmentData.HorizontalAlignment == Style.ExcelHorizontalAlignment.Left && cell.Worksheet.Cells[cell._fromRow, cell._fromCol + 1].Value != null) ||
                   (CellAlignmentData.HorizontalAlignment == Style.ExcelHorizontalAlignment.Right && cell.Worksheet.Cells[cell._fromRow, cell._fromCol - 1 <= 0 ? 1 : cell._fromCol-1 ].Value != null) )
                {
                    Clip = true;
                }
            }
        }

        private TtfFont CacheFont(PdfCellFontData fontData, Dictionary<string, PdfFontResource> fontResources, PdfPageSettings pageSettings)
        {
            string subfam = "Regular";
            if (FontData.Bold)
            {
                subfam = "Bold";
                if (FontData.Italic)
                {
                    subfam += " Italic";
                }
            }
            else if (FontData.Italic)
            {
                subfam = "Italic";
            }
            fontData.SubFamily = subfam;
            if (!fontResources.ContainsKey(FontData.FontName))
            {
                int label = 1;
                if (fontResources.Count > 0)
                {
                    label = fontResources.Last().Value.labelNumber + 1;
                }
                PdfFontResource fr = new PdfFontResource(FontData.FontName, subfam, label, pageSettings);
                fontResources.Add(FontData.FontName, fr);
                fontResources.Last().Value.fontData = PdfTextData.GetFontData(pageSettings, FontData.FontName, subfam);
            }
            return fontResources[FontData.FontName].fontData;
        }

        private Vector2 CalculatePosition(ExcelRangeBase cell, double cellX, double CellY, double cellWidth, double cellHeight, double textLength, double fontHeight)
        {
            double x = 0d;
            double y = 0d;
            switch (CellAlignmentData.HorizontalAlignment)
            {
                case Style.ExcelHorizontalAlignment.Fill:
                case Style.ExcelHorizontalAlignment.General:
                    if (double.TryParse(cell.Value.ToString(), out double value))
                    {
                        x = cellX + (cellWidth - textLength) - rightMargin;
                    }
                    else
                    {
                        x = cellX + rightMargin;
                    }
                    break;
                case Style.ExcelHorizontalAlignment.Left:
                    x = cellX + rightMargin;
                    break;
                case Style.ExcelHorizontalAlignment.Center:
                    x = (cellX + (cellWidth - textLength) / 2d);
                    break;
                case Style.ExcelHorizontalAlignment.Right:
                    x = cellX + (cellWidth - textLength) - rightMargin;
                    break;
            }
            switch (CellAlignmentData.VerticalAlignment)
            {
                case Style.ExcelVerticalAlignment.Top:
                    y = CellY + fontHeight + bottomMargin;
                    break;
                case Style.ExcelVerticalAlignment.Center:
                    y = CellY + (cellHeight - fontHeight) / 2d + fontHeight;
                    break;
                case Style.ExcelVerticalAlignment.Bottom:
                    y = ((CellY + (cellHeight - fontHeight)) + fontHeight) - bottomMargin;
                    break;
            }
            return new Vector2(x, y);
        }

        internal void AdjustClipping(List<PdfTransform> cells)
        {
            if (Clip)
            {
                var pcc = cells.Where(x => x.Name == this.Name).Where(x=> x is PdfCellLayout).ToList();
                if (pcc.Count > 0)
                {
                    Clipping = new PdfRect() { X = pcc[0].LocalPosition.X + rightMargin,
                                               Y = (pcc[0].LocalPosition.Y - pcc[0].Size.Y),
                                           Width = pcc[0].Size.X - (rightMargin*2),
                                          Height = pcc[0].Size.Y };
                }
            }
        }
    }
}

/*
 * check if text needs to be cut
 * if needs to be cut we can do the follwing solutions
 *      check for next cells and set Z to higher value than text
 */