using EPPlus.Export.Pdf.PdfLayout;
using EPPlus.Export.Pdf.PdfResources;
using EPPlus.Export.Pdf.PdfSettings;
using OfficeOpenXml;
using OfficeOpenXml.Interfaces.Fonts;
using OfficeOpenXml.Style;
using OfficeOpenXml.Style.HeaderFooterTextFormat;
using OfficeOpenXml.Style.Table;
using OfficeOpenXml.Table;
using System;
using System.Collections.Generic;
using System.Linq;

namespace EPPlus.Export.Pdf.PdfCatalog
{
    internal class PdfTextMap
    {
        public static PdfCellCollection SetTextMap(PdfPageSettings pageSettings, PdfDictionaries dictionaries, PdfWorksheet pdfSheet, PdfRange Range)
        {
            var Map = new PdfCellCollection(Range.Range._fromRow, Range.Range._toRow, Range.Range._fromCol, Range.Range._toCol);
            var worksheet = Range.Range.Worksheet;
            List<string> checkedMergedCells = new List<string>();
            //int addedColumns = Range.ExtendColumns ? AddColumnsForNonWrappedText(worksheet) : 0;
            for (int row = Range.Range._fromRow; row <= Range.Range._toRow; row++)
            {
                if (worksheet.Row(row).Hidden) continue;
                //var height = UnitConversion.ExcelRowHeightToPoints(worksheet.Row(row).Height);
                //x = 0d;
                for (int col = Range.Range._fromCol; col <= Range.Range._toCol /*+ addedColumns*/; col++)
                {
                    if (worksheet.Column(col).Hidden) continue;
                    //var width = UnitConversion.ExcelColumnWidthToPoints(worksheet.Column(col).Width, ZeroCharWidth);
                    var cell = worksheet.Cells[row, col];
                    //if (cell.Merge)
                    //{

                    //}
                    var tempMap = Map[row, col];

                    tempMap.CellStyle = GetFontStyle(cell);
                    tempMap.ContentAligmnet = GetContentAlignment(cell);
                    if (!cell.IsRichText) cell._rtc = new ExcelRichTextCollection(cell.Text, cell);
                    tempMap.TextFormats = GetTextFormats(pageSettings, dictionaries, cell._rtc, Map[row, col].CellStyle);

                    Map[row, col] = tempMap;
                    if (pageSettings.CommentsAndNotes != CommentsAndNotes.None)
                    {
                        if (cell.Comment != null && cell.ThreadedComment == null)
                        {
                            pdfSheet.CommentsAndNotesCollections.Add(cell.Address, new PdfCommentsAndNotes(cell.Comment));
                        }
                        if (cell.ThreadedComment != null)
                        {
                            pdfSheet.CommentsAndNotesCollections.Add(cell.Address, new PdfCommentsAndNotes(cell.ThreadedComment));
                            PdfCommentsAndNotes.HasThreadedComment = true;
                        }
                    }



                    //var cellStyle = new PdfCellStyle();
                    //GetFillStyles(cell, cellStyle);
                    //GetBorderStyles(cell, cellStyle);
                    //GetFontStyles(cell, cellStyle);
                    //PdfCellBorderLayout border = HandleEdgeBorders(cell, cellStyle, cell.Address, x, y, width, height);
                    //if (cell.Merge)
                    //{
                    //    HandleMergedCell(worksheet, pageSettings, dictionaries, cell, cellStyle, checkedMergedCells, x, y);
                    //}
                    //else
                    //{
                    //    HandleCell(pageSettings, dictionaries, cell, x, y, width, height, cellStyle);
                    //}
                    //if (border != null) border.InitEdgeBorders(cell);
                    //x += width;
                    //totalWidth = System.Math.Max(x, totalWidth);

                }
                //y -= height;
            }
            //HandleDrawings(worksheet);
            //Size = new Vector2(totalWidth, Math.Abs(y));
            return Map;
        }

        private static PdfCellStyle GetFontStyle(ExcelRangeBase cell)
        {
            var cellStyle = new PdfCellStyle();
            var tables = cell.Worksheet.Tables.GetIntersectingRanges(cell);
            if (tables.Count > 0)
            {
                var table = tables[0].Value;
                var range = table.Range;
                ExcelTableNamedStyle tableStyle;
                if (table.TableStyle == TableStyles.Custom)
                {
                    tableStyle = cell.Worksheet.Workbook.Styles.TableStyles[table.StyleName].As.TableStyle;
                }
                else
                {
                    var tmpNode = table.WorkSheet.Workbook.StylesXml.CreateElement("c:tableStyle");
                    tableStyle = new ExcelTableNamedStyle(cell.Worksheet.Workbook.Styles.NameSpaceManager, tmpNode, cell.Worksheet.Workbook.Styles);
                    tableStyle.SetFromTemplate((TableStyles)table.TableStyle);
                }
                int tableRow = cell._fromRow - range._fromRow;
                int tableCol = cell._fromCol - range._fromCol;
                var font = tableStyle.WholeTable.Style.Font;
                if (table.ShowHeader && tableRow == 0)
                {
                    if (tableStyle.HeaderRow.Style.Font.HasValue)
                    {
                        font = tableStyle.HeaderRow.Style.Font;
                    }
                    if (tableCol == 0 && table.ShowFirstColumn && tableStyle.FirstHeaderCell.Style.Font.HasValue)
                    {
                        font = tableStyle.FirstHeaderCell.Style.Font;
                    }
                    if (cell._fromCol == range._toCol && table.ShowLastColumn && tableStyle.LastHeaderCell.Style.Font.HasValue)
                    {
                        font = tableStyle.LastHeaderCell.Style.Font;
                    }
                }
                else if (table.ShowTotal && cell._fromRow == range._toRow)
                {
                    if (tableStyle.TotalRow.Style.Font.HasValue)
                    {
                        font = tableStyle.TotalRow.Style.Font;
                    }
                    if (tableCol == 0 && table.ShowFirstColumn && tableStyle.FirstTotalCell.Style.Font.HasValue)
                    {
                        font = tableStyle.FirstTotalCell.Style.Font;
                    }
                    if (cell._fromCol == range._toCol && table.ShowLastColumn && tableStyle.LastTotalCell.Style.Font.HasValue)
                    {
                        font = tableStyle.LastTotalCell.Style.Font;
                    }
                }
                else
                {
                    if (table.ShowColumnStripes && (tableCol & 1) == 0)
                    {
                        font = tableStyle.FirstColumnStripe.Style.Font;
                    }
                    if (table.ShowColumnStripes && tableStyle.SecondColumnStripe.Style.Border.Top.HasValue && (tableCol & 1) != 0)
                    {
                        font = tableStyle.SecondColumnStripe.Style.Font;
                    }
                    if (table.ShowRowStripes && tableStyle.FirstRowStripe.Style.Font.HasValue && (tableRow & 1) != 0)
                    {
                        font = tableStyle.FirstRowStripe.Style.Font;
                    }
                    if (table.ShowRowStripes && tableStyle.SecondRowStripe.Style.Font.HasValue && (tableRow & 1) == 0)
                    {
                        font = tableStyle.SecondRowStripe.Style.Font;
                    }
                    if (table.ShowLastColumn && tableStyle.LastColumn.Style.Font.HasValue && cell._fromCol == range._toCol)
                    {
                        font = tableStyle.LastColumn.Style.Font;
                    }
                    if (table.ShowFirstColumn && tableStyle.FirstColumn.Style.Font.HasValue && tableCol == range._toCol)
                    {
                        font = tableStyle.FirstColumn.Style.Font;
                    }
                }
                cellStyle.dxfFont = font;
            }
            return cellStyle;
        }


        private static PdfCellAlignmentData GetContentAlignment(ExcelRangeBase cell)
        {
            var contentAlignment = new PdfCellAlignmentData();
            contentAlignment.HorizontalAlignment = cell.Style.HorizontalAlignment;
            contentAlignment.VerticalAlignment = cell.Style.VerticalAlignment;
            contentAlignment.Indent = cell.Style.Indent;
            contentAlignment.WrapText = cell.Style.WrapText;
            contentAlignment.ShrinkToFit = cell.Style.ShrinkToFit;
            contentAlignment.TextRotation = (cell.Style.TextRotation >= 90) ? ((cell.Style.TextRotation == 255) ? 0 : 90 - cell.Style.TextRotation) : cell.Style.TextRotation;
            contentAlignment.IsVertical = cell.Style.TextRotation == 255 ? true : false;
            contentAlignment.TextDirection = cell.Style.ReadingOrder;
            return contentAlignment;
        }

        private static List<PdfTextFormat> GetTextFormats(PdfPageSettings pageSettings, PdfDictionaries dictionaries, ExcelRichTextCollection RichTextCollection, PdfCellStyle cellStyle = null)
        {
            var textFormats = new List<PdfTextFormat>();
            bool bold = false, italic = false, underline = false, strike = false;
            ExcelUnderLineType underLineType = ExcelUnderLineType.None;
            if (cellStyle != null && cellStyle.dxfFont != null)
            {
                bold = cellStyle.dxfFont.Bold != null ? (bool)cellStyle.dxfFont.Bold : false;
                italic = cellStyle.dxfFont.Italic != null ? (bool)cellStyle.dxfFont.Italic : false;
                strike = cellStyle.dxfFont.Strike != null ? (bool)cellStyle.dxfFont.Strike : false;
                underline = cellStyle.dxfFont.Underline != null;
                underLineType = cellStyle.dxfFont.Underline != null ? (ExcelUnderLineType)cellStyle.dxfFont.Underline : ExcelUnderLineType.None;
            }
            for (int i = 0; i < RichTextCollection.Count; i++)
            {
                var rt = RichTextCollection[i];
                var textFormat = new PdfTextFormat();
                textFormat.Text = rt.Text;
                textFormat.FontName = rt.FontName;
                textFormat.FontFamily = rt.Family;
                textFormat.FontSize = rt.Size;
                textFormat.Bold = rt.Bold || bold;
                textFormat.Italic = rt.Italic || italic;
                textFormat.Strike = rt.Strike || strike;
                textFormat.Underline = rt.UnderLine || underline;
                textFormat.UnderlineType = rt.UnderLineType == ExcelUnderLineType.None ? underLineType : rt.UnderLineType;
                textFormat.SuperScript = rt.VerticalAlign == ExcelVerticalAlignmentFont.Superscript;
                textFormat.SubScript = rt.VerticalAlign == ExcelVerticalAlignmentFont.Subscript;
                textFormat.FontColor = rt.Color;
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
                textFormats.Add(textFormat);
                dictionaries.AddFont(pageSettings, textFormat.FontName, textFormat.SubFamily, textFormat.Text);
            }
            return textFormats;
        }
        public static PdfHeaderFooter GetTextFormats(PdfPageSettings pageSettings, PdfDictionaries dictionaries, ExcelWorksheet ws, ExcelHeaderFooterTextCollection textCollection, HeaderFooterType type, HeaderFooterAlignment alignment, HeaderFooterSection section)
        {
            if (textCollection == null && textCollection.Count <= 0) return null;
            var textFormats = new List<PdfTextFormat>();
            bool containsPageNumber = false;
            for (int i = 1; i < textCollection.Count; i++)
            {
                var text = textCollection[i];
                var ns = ws.Workbook.Styles.GetNormalStyle();
                PdfTextFormat textFormat = new PdfTextFormat();
                textFormat.FontName = string.IsNullOrEmpty(text.FontName) ? ns.Style.Font.Name : text.FontName;
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
                    case ExcelHeaderFooterFormattingCodes.PageNumber:
                        textFormat.Text += "###";
                        containsPageNumber = true;
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
                textFormats.Add(textFormat);
                dictionaries.AddFont(pageSettings, textFormat.FontName, textFormat.SubFamily, textFormat.Text);
            }
            return new PdfHeaderFooter(textFormats, containsPageNumber, type, alignment, section);
        }
    }
}
