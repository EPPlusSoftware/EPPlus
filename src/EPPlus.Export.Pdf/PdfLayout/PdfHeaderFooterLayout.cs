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
using EPPlus.Export.Pdf.PdfResources;
using EPPlus.Export.Pdf.PdfSettings;
using EPPlus.Fonts.OpenType;
using EPPlus.Fonts.OpenType.Integration;
using EPPlus.Graphics;
using EPPlus.Graphics.Math;
using OfficeOpenXml;
using OfficeOpenXml.Interfaces.Drawing.Text;
using OfficeOpenXml.Interfaces.Fonts;
using OfficeOpenXml.Style;
using OfficeOpenXml.Style.HeaderFooterTextFormat;
using System;
using System.Collections.Generic;
using System.Linq;

namespace EPPlus.Export.Pdf.PdfLayout
{
    internal class PdfHeaderFooterLayout : Transform, ITextLayout
    {
        private List<PdfTextFormat> textFormats = new List<PdfTextFormat>();
        private double textLength { get; set; }
        private double textHeight { get; set; }
        private TextLayoutEngine textLayoutEngine;

        public List<PdfTextFormat> TextFormats { get => textFormats; set => textFormats = value; }
        public double TextLength { get => textLength; set => textLength = value; }
        public double TextHeight { get => textHeight; set => textHeight = value; }
        public TextLayoutEngine TextLayoutEngine { get => textLayoutEngine; set => textLayoutEngine = value; }

        public PdfHeaderFooterLayout(ExcelHeaderFooterTextCollection textCollection, ExcelWorksheet ws, PdfPageSettings pageSettings, PdfDictionaries dictionaries, int pageNumber, int totalPages)
        {
            for( int i=1; i < textCollection.Count; i++)
            {
                var text = textCollection[i];
                var ns = ws.Workbook.Styles.GetNormalStyle();
                PdfTextFormat textFormat = new PdfTextFormat();
                textFormat.FontName = string.IsNullOrEmpty( text.FontName) ? ns.Style.Font.Name : text.FontName;
                textFormat.FontSize = text.FontSize == null ? ns.Style.Font.Size : (double)text.FontSize;
                textFormat.Bold = text.Bold;
                textFormat.Italic = text.Italic;
                textFormat.Strike = text.Striketrough;
                textFormat.SubScript = text.SubScript;
                textFormat.SuperScript = text.SuperScript;
                textFormat.Underline = text.Underline;
                textFormat.UnderlineType = textFormat.Underline ? ExcelUnderLineType.Single : ExcelUnderLineType.None;
                if (text.DoubleUnderline) textFormat.Underline = true;
                textFormat.UnderlineType = text.DoubleUnderline ? ExcelUnderLineType.Double : textFormat.UnderlineType;
                textFormat.FontColor = text.Color;
                textFormat.SubFamily = FontSubFamily.Regular;
                if (textFormat.Bold)
                {
                    textFormat.SubFamily = FontSubFamily.Bold;
                    if (textFormat.Italic)
                    {
                        textFormat.SubFamily = FontSubFamily.BoldItalic;
                    }
                }
                else if (textFormat.Italic)
                {
                    textFormat.SubFamily = FontSubFamily.Italic;
                }

                switch (text.FormatCode)
                {
                    case ExcelHeaderFooterFormattingCodes.SheetName:
                        textFormat.Text += ws.Name;
                        break;
                    case ExcelHeaderFooterFormattingCodes.CurrentDate:
                        textFormat.Text += DateTime.Now.ToString($"yyyy-MM-dd");
                        break;
                    case ExcelHeaderFooterFormattingCodes.FileName:
                        textFormat.Text += ws._package.File.Name;
                        break;
                    case ExcelHeaderFooterFormattingCodes.NumberOfPages:
                        textFormat.Text += totalPages.ToString();
                        break;
                    case ExcelHeaderFooterFormattingCodes.PageNumber:
                        textFormat.Text += pageNumber;
                        break;
                    case ExcelHeaderFooterFormattingCodes.CurrentTime:
                        textFormat.Text += DateTime.Now.ToString("HH:mm");
                        break;
                    case ExcelHeaderFooterFormattingCodes.FilePath:
                        textFormat.Text += ws._package.File.Directory.FullName + "\\";
                        break;
                    default:
                        textFormat.Text += text.Text;
                        break;
                }
                PdfFontResource.GetFontResourceData(dictionaries.Fonts, pageSettings, textFormat);
                MeasurementFont font = new MeasurementFont();
                font.FontFamily = textFormat.FontName;
                font.Size = (float)textFormat.FontSize;
                font.Style = ((textFormat.Bold ? MeasurementFontStyles.Bold : 0) |
                             (textFormat.Italic ? MeasurementFontStyles.Italic : 0) |
                             (textFormat.Strike ? MeasurementFontStyles.Strikeout : 0) |
                             (textFormat.Underline ? MeasurementFontStyles.Underline : 0))
                             switch
                             {
                                 0 => MeasurementFontStyles.Regular,
                                 var s => s
                             };
                FontMeasurerTrueType fontMeasurerTrueType = new FontMeasurerTrueType();
                var result = fontMeasurerTrueType.MeasureText(textFormat.Text, font);
                textFormat.TextLength = result.Width;
                textLength += result.Width;
                textFormat.TextHeight = result.Height;
                textHeight = result.Height;
                textFormats.Add(textFormat);
                if (!dictionaries.Fonts.ContainsKey(textFormat.FullFontName))
                {
                    int label = 1;
                    if (dictionaries.Fonts.Count > 0)
                    {
                        label = dictionaries.Fonts.Last().Value.labelNumber + 1;
                    }
                    dictionaries.Fonts.Add(textFormat.FullFontName, new PdfFontResource(textFormat.FontName, textFormat.SubFamily, label, pageSettings));
                }
                var manger = dictionaries.Fonts[textFormat.FullFontName].fontSubsetManager;
                manger.AddText(textFormat.Text);
            }
        }

        public void AdjustPositionByTextLength(char rc, char hf)
        {
            if (rc == 'r')
            {
                LocalPosition = new Vector2(LocalPosition.X - textLength, LocalPosition.Y);
            }
            else if (rc == 'c')
            {
                LocalPosition = new Vector2(LocalPosition.X - (textLength/2d), LocalPosition.Y);
            }
            if (hf == 'h')
            {
                LocalPosition = new Vector2(LocalPosition.X, LocalPosition.Y - textHeight);
            }
        }

    }
}
