using FontLab1;
using OfficeOpenXml.PDF.Math;
using OfficeOpenXml.PDF.PdfGraphics;
using OfficeOpenXml.PDF.Pdfhelpers;
using OfficeOpenXml.PDF.PdfResources;
using OfficeOpenXml.PDF.PdfSettings;
using System.Collections.Generic;
using System.Linq;

namespace OfficeOpenXml.PDF.PdfLayout
{
    internal class PdfCellContentLayout : PdfTransform, ILayout
    {
        public PdfCellFontData FontData;
        public PdfCellAlignmentData CellAlignmentData;
        public bool Clip;
        public PdfRect Clipping;

        private double bottomMargin = 3.4d; //Guessed number
        private double rightMargin = 1.0d; //I guessed this one too..

        public PdfCellContentLayout(ExcelRangeBase cell, PdfPageSettings pageSettings, double x, double y, double width, double height, double scaleX = 1, double scaleY = 1, double rotation = 0, PdfTransform parent = null, PdfDictionaries dictionaries = null)
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
            FontData.Text = cell.Text;
            CellAlignmentData = new PdfCellAlignmentData();
            CellAlignmentData.HorizontalAlignment = cell.Style.HorizontalAlignment;
            CellAlignmentData.VerticalAlignment = cell.Style.VerticalAlignment;
            CellAlignmentData.Indent = cell.Style.Indent;
            CellAlignmentData.WrapText = cell.Style.WrapText;
            CellAlignmentData.ShrinkToFit = cell.Style.ShrinkToFit; //Need to fix Transform issues and then implement a method that sets scale on the text object.
            CellAlignmentData.TextRotation = cell.Style.TextRotation; //EPPlus does probably not calculate cell width and height after setting rotation on text. So before we make pdf we need to calculate cell width and height based on text rotation
            CellAlignmentData.TextDirection = cell.Style.ReadingOrder;
            var ttfont = GetFontResourceData(dictionaries.Fonts, pageSettings);
            double textLength = PdfTextData.MeasureText(FontData.Text, FontData.FontSize, ttfont);
            double fontHeight = PdfTextData.MeasureFontHeight(ttfont, FontData.FontSize);
            LocalPosition = CalculatePosition(cell, x, y, width, height, textLength, fontHeight);
            Size = new Vector2((x + width) - LocalPosition.X, (y + height) - LocalPosition.Y);
            CheckClipping(cell, textLength, width);
        }

        //Get font data from fontResources. If font does not exsist, add it to fontResources.
        private TtfFont GetFontResourceData(Dictionary<string, PdfFontResource> fontResources, PdfPageSettings pageSettings)
        {
            if (!fontResources.ContainsKey(FontData.Font))
            {
                int label = 1;
                if (fontResources.Count > 0)
                {
                    label = fontResources.Last().Value.labelNumber + 1;
                }
                PdfFontResource fr = new PdfFontResource(FontData.FontName, FontData.SubFamily, label, pageSettings);
                fontResources.Add(FontData.Font, fr);
                fontResources.Last().Value.fontData = PdfTextData.GetFontData(pageSettings, FontData.FontName, FontData.SubFamily);
            }
            return fontResources[FontData.Font].fontData;
        }

        //Calculate text position based on alignment and cell size.
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

        //Check if clipping is needed.
        private void CheckClipping(ExcelRangeBase cell, double textLength, double width)
        {
            if (textLength >= width)
            {
                if (CellAlignmentData.HorizontalAlignment == Style.ExcelHorizontalAlignment.Fill ||
                   (CellAlignmentData.HorizontalAlignment == Style.ExcelHorizontalAlignment.Left  && cell.Worksheet.Cells[cell._fromRow, cell._fromCol + 1].Value != null) ||
                   (CellAlignmentData.HorizontalAlignment == Style.ExcelHorizontalAlignment.Right && cell.Worksheet.Cells[cell._fromRow, ((cell._fromCol - 1 <= 0) ? (1) : (cell._fromCol - 1))].Value != null))
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
                var pcc = cells.Where(x => x.Name == this.Name).Where(x=> x is PdfCellLayout).ToList();
                if (pcc.Count > 0)
                {
                    Clipping = new PdfRect()
                    {
                        X = pcc[0].LocalPosition.X + rightMargin,
                        Y = (pcc[0].LocalPosition.Y),
                        Width = pcc[0].Size.X - (rightMargin * 2),
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
