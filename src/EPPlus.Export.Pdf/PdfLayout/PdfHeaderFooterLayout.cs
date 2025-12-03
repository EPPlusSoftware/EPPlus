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
using OfficeOpenXml;
using OfficeOpenXml.Style.HeaderFooterTextFormat;
using System;
using System.Collections.Generic;
using System.Linq;


namespace EPPlus.Export.Pdf.PdfLayout
{
    internal class PdfHeaderFooterLayout : Transform
    {
        public PdfCellTextLine textLine = new PdfCellTextLine();

        public PdfHeaderFooterLayout(ExcelHeaderFooterTextCollection textCollection, ExcelWorksheet ws, PdfPageSettings settings, PdfDictionaries dictionaries, int pageNumber, int totalPages)
        {
            foreach (var text in textCollection)
            {
                PdfCellTextItem textItem = new PdfCellTextItem();
                textItem.FontName = text.FontName;
                textItem.FontSize = (double)text.FontSize;
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
                textLine.TextItemCollection.Add(textItem);
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

    }
}
