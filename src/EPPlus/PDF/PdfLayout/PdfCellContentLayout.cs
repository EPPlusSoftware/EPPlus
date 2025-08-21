using FontLab1;
using OfficeOpenXml.PDF.Math;
using OfficeOpenXml.PDF.PdfFontData;
using OfficeOpenXml.PDF.PdfGraphics;
using OfficeOpenXml.PDF.Pdfhelpers;
using OfficeOpenXml.PDF.PdfSettings;
using System.Collections.Generic;
using System.Linq;

namespace OfficeOpenXml.PDF.PdfLayout
{

    internal class PdfCellContentLayout : PdfTransform
    {
        public PdfCellFontData FontData;
        public PdfCellAlignmentData CellAlignmentData;

        public PdfCellContentLayout(ExcelRangeBase cell, PdfPageSettings pageSettings, double x, double y, double width, double height, double scaleX = 1, double scaleY = 1, double rotation = 0, PdfTransform parent = null, Dictionary<string, PdfFontResource> fontResources = null)
            : base(x, y, 0, 0, scaleX, scaleY, rotation, parent)
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
            CellAlignmentData.ShrinkToFit = cell.Style.ShrinkToFit;
            CellAlignmentData.TextRotation = cell.Style.TextRotation;
            CellAlignmentData.TextDirection = cell.Style.ReadingOrder;

            var ttfont = CacheFont(FontData, fontResources, pageSettings);
            double textLength = PdfTextData.MeasureText(FontData.Text, FontData.FontSize, ttfont);
            double fontHeight = PdfTextData.MeasureFontHeight(ttfont, FontData.FontSize);
            LocalPosition = CalculatePosition(cell, x, y, width, height, textLength, fontHeight);
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
            var bottomMargin = 20d; //Guessed number
            var rightMargin = 0.8d; //I guessed this one too..
            switch (CellAlignmentData.HorizontalAlignment)
            {
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
                    y = CellY + fontHeight;
                    break;
                case Style.ExcelVerticalAlignment.Center:
                    y = CellY + (cellHeight - fontHeight) / 2d;
                    break;
                case Style.ExcelVerticalAlignment.Bottom:
                    y = (CellY + (cellHeight - fontHeight)) + fontHeight;
                    break;

            }
            return new Vector2(x, y);
        }

    }
}

/*
 * check if text needs to be cut
 * if needs to be cut we can do the follwing solutions
 *      check for next cells and set Z to higher value than text
 
 */