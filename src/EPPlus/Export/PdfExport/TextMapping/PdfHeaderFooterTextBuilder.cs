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
using EPPlus.Export.Pdf.Enums;
using EPPlus.Export.Pdf.Layout;
using EPPlus.Export.Pdf.Resources;
using EPPlus.Export.Pdf.Settings;
using EPPlus.Fonts.OpenType.Integration;
using EPPlus.Fonts.OpenType.Integration.DataHolders;
using OfficeOpenXml.Drawing.Interfaces;
using OfficeOpenXml.Style.HeaderFooterTextFormat;
using System;
using System.Collections.Generic;

namespace OfficeOpenXml.Export.PdfExport.TextMapping
{
    internal static class PdfHeaderFooterTextBuilder
    {
        internal static PdfHeaderFooter GetTextFormats(PdfRenderSettings pageSettings, PdfDictionaries dictionaries, ExcelWorksheet ws, ExcelHeaderFooterTextCollection textCollection, HeaderFooterType type, HeaderFooterAlignment alignment, HeaderFooterSection section)
        {
            if (textCollection == null || textCollection.Count <= 1) return null;
            var ns = ws.Workbook.Styles.GetNormalStyle();
            var textFragments = new List<TextFragment>();
            List<int> NumberOfPagesIndexes = new List<int>();
            List<int> PageNumberIndexes = new List<int>();
            byte[] imageBytes = null;
            double imageWidth = 0, imageHeight = 0;
            int imageFragmentIndex = -1;
            for (int i = 1; i < textCollection.Count; i++)
            {
                var hf = textCollection[i];
                if (hf.FormatCode == ExcelHeaderFooterFormattingCodes.Image)
                {
                    var pic = textCollection.Picture;
                    var picUri = pic == null ? null : ((IPictureContainer)pic).UriPic;
                    var picBytes = picUri == null ? null : ws._package.PictureStore.GetImageBytes(picUri);
                    if (picBytes != null)
                    {
                        imageBytes = picBytes;
                        imageWidth = pic.Width;
                        imageHeight = pic.Height;
                        imageFragmentIndex = textFragments.Count;
                    }
                    continue;
                }
                var textFrag = new TextFragment();
                textFrag.Font = new RichTextFormatSimple();
                textFrag.Font.Family = string.IsNullOrEmpty(hf.FontName) ? ns.Style.Font.Name : hf.FontName;
                textFrag.Font.Size = hf.FontSize == null ? ns.Style.Font.Size : (float)hf.FontSize;
                textFrag.RichTextOptions.Bold = hf.Bold;
                textFrag.RichTextOptions.Italic = hf.Italic;
                //underline
                //none   : 12
                //single : 13
                //Double : 4
                //accouting does not exsist
                textFrag.RichTextOptions.UnderlineType = 12;
                textFrag.RichTextOptions.UnderlineType = hf.Underline ? 13 : textFrag.RichTextOptions.UnderlineType;
                textFrag.RichTextOptions.UnderlineType = hf.DoubleUnderline ? 4 : textFrag.RichTextOptions.UnderlineType;
                textFrag.RichTextOptions.StrikeType = hf.Striketrough ? 2 : 1;
                textFrag.RichTextOptions.FontColor = hf.Color;
                textFrag.Font.SubFamily = PdfTextFragmentBuilder.ComputeFontStyle(textFrag);
                var text = string.Empty;
                switch (hf.FormatCode)
                {
                    case ExcelHeaderFooterFormattingCodes.SheetName:
                        text += ws.Name;
                        break;
                    case ExcelHeaderFooterFormattingCodes.CurrentDate:
                        text += DateTime.Now.ToString($"yyyy-MM-dd");
                        break;
                    case ExcelHeaderFooterFormattingCodes.FileName:
                        text += ws._package.File.Name;
                        break;
                    case ExcelHeaderFooterFormattingCodes.NumberOfPages:
                        text += "000";
                        NumberOfPagesIndexes.Add(i - 1);
                        break;
                    case ExcelHeaderFooterFormattingCodes.PageNumber:
                        text += "000";
                        PageNumberIndexes.Add(i - 1);
                        break;
                    case ExcelHeaderFooterFormattingCodes.CurrentTime:
                        text += DateTime.Now.ToString("HH:mm");
                        break;
                    case ExcelHeaderFooterFormattingCodes.FilePath:
                        text += ws._package.File.Directory.FullName + "\\";
                        break;
                    default:
                        text += hf.Text;
                        break;
                }
                textFrag.Text = text;
                textFragments.Add(textFrag);
                dictionaries.AddFont(pageSettings, textFrag.Font.Family, textFrag.Font.SubFamily, textFrag.Text);
                if (NumberOfPagesIndexes.Count > 0 || PageNumberIndexes.Count > 0) dictionaries.AddFont(pageSettings, textFrag.Font.Family, textFrag.Font.SubFamily, "1234567890");
            }
            var result = new PdfHeaderFooter(textFragments, PageNumberIndexes, NumberOfPagesIndexes, type, alignment, section);
            result.ImageBytes = imageBytes;
            result.ImageWidth = imageWidth;
            result.ImageHeight = imageHeight;
            result.ImageFragmentIndex = imageFragmentIndex;
            return result;
        }

        internal static PdfCellAlignmentData GetAlignmentData(PdfHeaderFooter headerFooter)
        {
            var contentAlignment = new PdfCellAlignmentData();
            switch (headerFooter.Alignment)
            {
                case HeaderFooterAlignment.Left:
                    contentAlignment.HorizontalAlignment = (EPPlus.Export.Pdf.Enums.ExcelHorizontalAlignment)ExcelHorizontalAlignment.Left;
                    break;
                case HeaderFooterAlignment.Center:
                    contentAlignment.HorizontalAlignment = (EPPlus.Export.Pdf.Enums.ExcelHorizontalAlignment)ExcelHorizontalAlignment.Center;
                    break;
                case HeaderFooterAlignment.Right:
                    contentAlignment.HorizontalAlignment = (EPPlus.Export.Pdf.Enums.ExcelHorizontalAlignment)ExcelHorizontalAlignment.Right;
                    break;
            }
            return contentAlignment;
        }
    }
}
