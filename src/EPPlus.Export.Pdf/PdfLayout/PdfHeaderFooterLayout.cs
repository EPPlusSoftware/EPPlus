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
using OfficeOpenXml.Style.HeaderFooterTextFormat;
using System;
using System.Collections.Generic;
using System.Linq;

namespace EPPlus.Export.Pdf.PdfLayout
{
    internal class PdfHeaderFooterLayout : Transform
    {
        public PdfCellWord textLine = new PdfCellWord();

        public PdfHeaderFooterLayout(ExcelHeaderFooterTextCollection textCollection, ExcelWorksheet ws, PdfPageSettings settings, PdfDictionaries dictionaries, int pageNumber, int totalPages)
        {
            for( int i=1; i < textCollection.Count; i++)
            {
                var text = textCollection[i];
                var ns = ws.Workbook.Styles.GetNormalStyle();
                PdfCellTextItem textItem = new PdfCellTextItem();
                textItem.FontName = string.IsNullOrEmpty( text.FontName) ? ns.Style.Font.Name : text.FontName;
                textItem.FontSize = text.FontSize == null ? ns.Style.Font.Size : (double)text.FontSize;
                textItem.Bold = text.Bold;
                textItem.Italic = text.Italic;
                textItem.Strike = text.Striketrough;
                textItem.SubScript = text.SubScript;
                textItem.SuperScript = text.SuperScript;
                textItem.Underline = text.Underline;
                textItem.FontColor = text.Color;
                switch (text.FormatCode)
                {
                    case ExcelHeaderFooterFormattingCodes.SheetName:
                        textItem.Text += ws.Name;
                        break;
                    case ExcelHeaderFooterFormattingCodes.CurrentDate:
                        textItem.Text += DateTime.Now.ToString("yyyy-MM-dd");
                        break;
                    case ExcelHeaderFooterFormattingCodes.FileName:
                        textItem.Text += ws._package.File.Name;
                        break;
                    case ExcelHeaderFooterFormattingCodes.NumberOfPages:
                        textItem.Text += totalPages.ToString();
                        break;
                    case ExcelHeaderFooterFormattingCodes.PageNumber:
                        textItem.Text += pageNumber;
                        break;
                    case ExcelHeaderFooterFormattingCodes.CurrentTime:
                        textItem.Text += DateTime.Now.ToString("HH:mm");
                        break;
                    case ExcelHeaderFooterFormattingCodes.FilePath:
                        textItem.Text += ws._package.File.FullName;
                        break;
                    default:
                        textItem.Text += text.Text;
                        break;
                }
                GetFontResourceData(dictionaries.Fonts, settings, textItem);
                MeasurementFont font = new MeasurementFont();
                font.FontFamily = textItem.FontName;
                font.Size = (float)textItem.FontSize;
                font.Style = ((textItem.Bold ? MeasurementFontStyles.Bold : 0) |
                             (textItem.Italic ? MeasurementFontStyles.Italic : 0) |
                             (textItem.Strike ? MeasurementFontStyles.Strikeout : 0) |
                             (textItem.Underline ? MeasurementFontStyles.Underline : 0))
                             switch
                             {
                                 0 => MeasurementFontStyles.Regular,
                                 var s => s
                             };
                FontMeasurerTrueType fontMeasurerTrueType = new FontMeasurerTrueType();
                var result = fontMeasurerTrueType.MeasureText(textItem.Text, font);
                textItem.TextLength = result.Width;
                textItem.FontHeight = result.FontHeight;
                textItem.LineHeight = result.Height;
                textLine.Characters.Add(textItem);
            }
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

        public void AdjustPositionByTextLength(char rc, char hf)
        {
            if (rc == 'r')
            {
                LocalPosition = new Vector2(LocalPosition.X - textLine.TextLength, LocalPosition.Y);
            }
            else if (rc == 'c')
            {
                LocalPosition = new Vector2(LocalPosition.X - (textLine.TextLength/2d), LocalPosition.Y);
            }
            if (hf == 'h')
            {
                LocalPosition = new Vector2(LocalPosition.X, LocalPosition.Y - textLine.LineHeight);
            }
        }

    }
}
