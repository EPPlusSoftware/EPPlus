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
using EPPlus.Export.Pdf.Layout;
using EPPlus.Export.Pdf.Resources;
using EPPlus.Export.Pdf.Settings;
using EPPlus.Fonts.OpenType.Integration;
using EPPlus.Fonts.OpenType.Integration.DataHolders;
using EPPlus.Graphics.Units;
using OfficeOpenXml.Export.PdfExport.Data;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Text;
using OfficeOpenXml.Interfaces.Fonts;
using OfficeOpenXml.Style;
using OfficeOpenXml.Style.Dxf;
using OfficeOpenXml.Style.HeaderFooterTextFormat;
using OfficeOpenXml.Style.Table;
using OfficeOpenXml.Table;
using System;
using System.Collections.Generic;
using System.Linq;

namespace OfficeOpenXml.Export.PdfExport.TextMapping
{
    internal class PdfTextMap
    {
        public static PdfCellCollection SetTextMap(PdfPageSettings pageSettings, PdfDictionaries dictionaries, PdfWorksheet pdfSheet, ref PdfRange pdfRange)
        {
            var tableStyleCache = new Dictionary<ExcelTable, ExcelTableNamedStyle>();
            var Range = pdfRange;
            var worksheet = Range.Range.Worksheet;
            var ZeroCharWidth = pdfSheet.ZeroCharWidth = PdfWorksheet.GetThemeFont0Width(worksheet);
            int addedColumns = Range.ExtendColumns ? AddColumnsForNonWrappedText(pageSettings, worksheet, pdfSheet) : 0;
            var Map = new PdfCellCollection(Range.Range._fromRow, Range.Range._toRow, Range.Range._fromCol, Range.Range._toCol + addedColumns);
            pdfSheet.ToRow = pdfSheet.ToRow < Range.Range._toRow ? Range.Range._toRow : pdfSheet.ToRow;
            bool firstColumnRun = true;
            List<string> checkedMergedCells = new List<string>();
            for (int row = Range.Range._fromRow; row <= Range.Range._toRow; row++)
            {
                var hiddenRow = worksheet.Row(row).Hidden;
                var r = (RowInternal)worksheet.GetValueInner(row, 0);
                bool usesDefaultValue = false;
                double height = 0;
                if (r == null || r.Height < 0)
                {
                    usesDefaultValue = true;
                    height = worksheet.DefaultRowHeight;
                }
                else
                {
                    height = UnitConversion.ExcelRowHeightToPoints(r.Height);
                }
                Range.TotalHeight += hiddenRow ? 0d : height;
                Range.RowHeights.Add(new RowHeight { Height = hiddenRow ? 0d : height, UsesDefaultValue = usesDefaultValue });
                for (int col = Range.Range._fromCol; col <= Range.Range._toCol + addedColumns; col++)
                {
                    var hiddenCol = worksheet.Column(col).Hidden;
                    var width = UnitConversion.ExcelColumnWidthToPoints(worksheet.Column(col).Width, ZeroCharWidth);
                    if (firstColumnRun)
                    {
                        Range.TotalWidth += hiddenCol ? 0d : width;
                        Range.ColWidths.Add(hiddenCol ? 0d : width);
                    }
                    var tempMap = new PdfCell();
                    tempMap.Hidden = hiddenRow || hiddenCol;
                    tempMap.ColumnWidth = tempMap.Width = hiddenCol ? 0d : width;
                    var cell = worksheet.Cells[row, col];
                    tempMap.Name = cell.Address;
                    if (cell.Merge)
                    {
                        HandleMergedCell(pageSettings, dictionaries, cell, checkedMergedCells, Map, tempMap, pdfSheet.ZeroCharWidth, tableStyleCache);
                    }
                    var cellStyle = new PdfCellStyle();
                    GetBorderStyles(cell, cellStyle, tempMap, tableStyleCache);
                    if (tempMap.Main == null)
                    {
                        GetFillStyles(cell, cellStyle, tableStyleCache);
                        GetFontStyle(cell, cellStyle, tableStyleCache);
                        tempMap.ContentAligmnet = GetContentAlignment(cell);
                        if (!string.IsNullOrEmpty(cell.Text))
                        {
                            tempMap.Text = cell.Text;
                            tempMap.TextFragments = GetTextFragments(pageSettings, dictionaries, cell, cellStyle);
                        }
                    }
                    tempMap.CellStyle = cellStyle;
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
                }
                firstColumnRun = false;
            }
            worksheet.ConditionalFormatting.ClearTempExportCacheForAllCFs();
            pdfRange = Range;
            ReconcileSharedBorders(Map);
            return Map;
        }

        private static void HandleMergedCell(PdfPageSettings pageSettings, PdfDictionaries dictionaries, ExcelRange cell, List<string> checkedMergedCells, PdfCellCollection map, PdfCell tempMap, double ZeroCharWidth, Dictionary<ExcelTable, ExcelTableNamedStyle> tableStyleCache)
        {
            var worksheet = cell.Worksheet;
            string mergeAddress = worksheet.MergedCells[cell.Start.Row, cell.Start.Column];
            ExcelAddressBase address = new ExcelAddressBase(mergeAddress);
            if (!checkedMergedCells.Contains(mergeAddress))
            {
                double totalWidth = 0, totalHeight = 0;
                for (int k = address._fromRow; k <= address._toRow; k++)
                {
                    if (worksheet.Row(k).Hidden) continue;
                    totalHeight += UnitConversion.ExcelRowHeightToPoints(worksheet.Row(k).Height);
                }
                for (int l = address._fromCol; l <= address._toCol; l++)
                {
                    if (worksheet.Column(l).Hidden) continue;
                    totalWidth += UnitConversion.ExcelColumnWidthToPoints(worksheet.Column(l).Width, ZeroCharWidth);
                }
                checkedMergedCells.Add(mergeAddress);
                tempMap.Width = totalWidth;
                tempMap.Height = totalHeight;
                tempMap.Main = null;
            }
            else
            {
                tempMap.Main = map[address._fromRow, address._fromCol];
                if (tempMap.Main == null)
                {
                    var main = worksheet.Cells[address._fromRow, address._fromCol];
                    PdfCell mainCell = new PdfCell();
                    var cellStyle = new PdfCellStyle();
                    GetBorderStyles(main, cellStyle, mainCell, tableStyleCache);
                    GetFillStyles(main, cellStyle, tableStyleCache);
                    GetFontStyle(main, cellStyle, tableStyleCache);
                    mainCell.ContentAligmnet = GetContentAlignment(main);
                    if (!string.IsNullOrEmpty(main.Text))
                    {
                        mainCell.Text = main.Text;
                        mainCell.TextFragments = GetTextFragments(pageSettings, dictionaries, main, cellStyle);
                    }
                    mainCell.CellStyle = cellStyle;
                    tempMap.Main = mainCell;
                }
            }
            tempMap.MergedAddress = address;
            tempMap.Name = tempMap.Name + " ; " + address.ToString();
            tempMap.Merged = true;
        }

        private static void GetFillStyles(ExcelRangeBase cell, PdfCellStyle cellStyle, Dictionary<ExcelTable, ExcelTableNamedStyle> tableStyleCache)
        {
            if (cell.Style.Fill.IsEmpty())
            {
                //Conditional Formating
                var cf = cell.ConditionalFormatting.GetConditionalFormattings();
                if (cf != null && cf.Count > 0)
                {
                    // Sort ascending — priority 1 beats priority 2, etc.
                    var ordered = cf.OrderBy(r => r.Priority);
                    foreach (var rule in ordered)
                    {
                        // Use the core per-rule evaluator (the same one the HTML exporter
                        // calls). It correctly handles every rule type — comparisons, text,
                        // blanks/errors, top/bottom, above/below average, duplicate/unique,
                        // time periods and formula expressions — including the range-wide
                        // aggregates, which the previous hand-rolled evaluator stubbed out.
                        if (!rule.ShouldApplyToCell(cell))
                        {
                            if (rule.StopIfTrue) break; // higher-priority rule fired but had no fill — stop anyway
                            continue;
                        }

                        if (rule.Style?.Fill != null && rule.Style.Fill.HasValue)
                        {
                            cellStyle.dxfFill = rule.Style.Fill;
                            // xfFill must be non-null for the downstream dxf path
                            // (PdfCellLayout checks xfFill.IsEmpty()); the cell's own fill
                            // is empty here, which is exactly what selects the dxf fill.
                            cellStyle.xfFill = cell.Style.Fill;
                            return; // CF fill wins — skip table and xf entirely
                        }

                        if (rule.StopIfTrue) break;
                    }
                }
                //Table
                var tables = cell.Worksheet.Tables.GetIntersectingRanges(cell);
                if (tables.Count > 0)
                {
                    var table = tables[0].Value;
                    var range = table.Range;
                    int tableRow = 0;
                    int tableCol = 0;
                    ExcelTableNamedStyle tableStyle = GetTableStyle(table, tableStyleCache);
                    tableRow = cell._fromRow - range._fromRow;
                    tableCol = cell._fromCol - range._fromCol;
                    if (table.ShowHeader && tableRow == 0)
                    {
                        cellStyle.dxfFill = tableStyle.HeaderRow.Style.Fill;
                    }
                    if (table.ShowHeader && tableRow == 0)
                    {
                        cellStyle.dxfFill = tableStyle.HeaderRow.Style.Fill;
                    }
                    else if (table.ShowTotal && range._toRow == cell._fromRow)
                    {
                        cellStyle.dxfFill = tableStyle.TotalRow.Style.Fill;
                    }
                    else if (table.ShowFirstColumn && tableCol == 0)
                    {
                        cellStyle.dxfFill = tableStyle.FirstColumn.Style.Fill;
                    }
                    else if (table.ShowLastColumn && range._toCol == cell._fromCol)
                    {
                        cellStyle.dxfFill = tableStyle.LastColumn.Style.Fill;
                    }
                    else if (table.ShowRowStripes)
                    {
                        var stripe = (tableRow & 1) == 0
                            ? tableStyle.SecondRowStripe.Style.Fill
                            : tableStyle.FirstRowStripe.Style.Fill;
                        cellStyle.dxfFill = FillIsPaintable(stripe)
                            ? stripe
                            : tableStyle.WholeTable.Style.Fill;
                    }
                    else if (table.ShowColumnStripes)
                    {
                        var stripe = (tableCol & 1) != 0
                            ? tableStyle.SecondColumnStripe.Style.Fill
                            : tableStyle.FirstColumnStripe.Style.Fill;
                        cellStyle.dxfFill = FillIsPaintable(stripe)
                            ? stripe
                            : tableStyle.WholeTable.Style.Fill;
                    }
                }
            }
            cellStyle.xfFill = cell.Style.Fill;
        }

        private static void GetBorderStyles(ExcelRangeBase cell, PdfCellStyle cellStyle, PdfCell pcell, Dictionary<ExcelTable, ExcelTableNamedStyle> tableStyleCache)
        {
            if (cell != null)
            {
                cellStyle.xfTop = cell.Style.Border.Top;
                cellStyle.xfBottom = cell.Style.Border.Bottom;
                cellStyle.xfLeft = cell.Style.Border.Left;
                cellStyle.xfRight = cell.Style.Border.Right;
                if (pcell.Main == null)
                {
                    cellStyle.Diagonal = cell.Style.Border.Diagonal;
                    cellStyle.DiagonalUp = cell.Style.Border.DiagonalUp;
                    cellStyle.DiagonalDown = cell.Style.Border.DiagonalDown;
                }
                else
                {
                    cellStyle.Diagonal = cell.Style.Border.Diagonal;
                    cellStyle.DiagonalUp = false;
                    cellStyle.DiagonalDown = false;
                }
                var tables = cell.Worksheet.Tables.GetIntersectingRanges(cell);
                if (tables.Count > 0)
                {
                    var table = tables[0].Value;
                    ExcelTableNamedStyle tableStyle = GetTableStyle(table, tableStyleCache);
                    if (tableStyle != null)
                    {
                        cellStyle.dxfTop = GetTopBorderItem(cell, cellStyle.xfTop, table, tableStyle, out int topOrder);
                        cellStyle.dxfTopElementOrder = topOrder;
                        cellStyle.dxfBottom = GetBottomBorderItem(cell, cellStyle.xfBottom, table, tableStyle, out int botOrder);
                        cellStyle.dxfBottomElementOrder = botOrder;
                        cellStyle.dxfLeft = GetLeftBorderItem(cell, cellStyle.xfLeft, table, tableStyle, out int leftOrder);
                        cellStyle.dxfLeftElementOrder = leftOrder;
                        cellStyle.dxfRight = GetRightBorderItem(cell, cellStyle.xfRight, table, tableStyle, out int rightOrder);
                        cellStyle.dxfRightElementOrder = rightOrder;
                    }
                }
                var cfBorder = GetConditionalFormattingBorder(cell);
                if (cfBorder != null)
                {
                    if (cfBorder.Top != null && cfBorder.Top.HasValue)
                    {
                        cellStyle.dxfTop = cfBorder.Top;
                        cellStyle.dxfTopElementOrder = TableEdgeOrder.ConditionalFormat;
                    }
                    if (cfBorder.Bottom != null && cfBorder.Bottom.HasValue)
                    {
                        cellStyle.dxfBottom = cfBorder.Bottom;
                        cellStyle.dxfBottomElementOrder = TableEdgeOrder.ConditionalFormat;
                    }
                    if (cfBorder.Left != null && cfBorder.Left.HasValue)
                    {
                        cellStyle.dxfLeft = cfBorder.Left;
                        cellStyle.dxfLeftElementOrder = TableEdgeOrder.ConditionalFormat;
                    }
                    if (cfBorder.Right != null && cfBorder.Right.HasValue)
                    {
                        cellStyle.dxfRight = cfBorder.Right;
                        cellStyle.dxfRightElementOrder = TableEdgeOrder.ConditionalFormat;
                    }
                }
            }
        }

        private static ExcelDxfBorderBase GetConditionalFormattingBorder(ExcelRangeBase cell)
        {
            var cf = cell.ConditionalFormatting.GetConditionalFormattings();
            if (cf != null && cf.Count > 0)
            {
                var ordered = cf.OrderBy(r => r.Priority);
                foreach (var rule in ordered)
                {
                    if (!rule.ShouldApplyToCell(cell))
                    {
                        if (rule.StopIfTrue) break;
                        continue;
                    }
                    if (rule.Style?.Border != null && rule.Style.Border.HasValue)
                    {
                        return rule.Style.Border;
                    }
                    if (rule.StopIfTrue) break;
                }
            }
            return null;
        }

        private static void GetTableRegionOverride(ExcelRangeBase cell, ExcelTable table,
            out ExcelDxfStyle colStyle, out ExcelDxfStyle tblStyle)
        {
            colStyle = null;
            tblStyle = null;
            var range = table.Range;
            int tableRow = cell._fromRow - range._fromRow;
            int tableCol = cell._fromCol - range._fromCol;
            ExcelTableColumn column = (tableCol >= 0 && tableCol < table.Columns.Count) ? table.Columns[tableCol] : null;

            string attr;
            ExcelDxfStyle tblCandidate, colCandidate;
            if (table.ShowHeader && tableRow == 0)
            { attr = "headerRowDxfId"; tblCandidate = table.HeaderRowStyle; colCandidate = column?.HeaderRowStyle; }
            else if (table.ShowTotal && cell._fromRow == range._toRow)
            { attr = "totalsRowDxfId"; tblCandidate = table.TotalsRowStyle; colCandidate = column?.TotalsRowStyle; }
            else
            { attr = "dataDxfId"; tblCandidate = table.DataStyle; colCandidate = column?.DataStyle; }

            var tableNode = table.TableXml?.DocumentElement;
            if (tableNode == null) return;   // no raw XML: treat as no override (leaves the style element in place)

            if (tableNode.Attributes?[attr] != null) tblStyle = tblCandidate;

            if (column != null)
            {
                var colNode = GetTableColumnNode(tableNode, tableCol);
                if (colNode?.Attributes?[attr] != null) colStyle = colCandidate;
            }
        }

        private static System.Xml.XmlNode GetTableColumnNode(System.Xml.XmlNode tableNode, int index)
        {
            foreach (System.Xml.XmlNode child in tableNode.ChildNodes)
            {
                if (child.LocalName != "tableColumns") continue;
                int i = 0;
                foreach (System.Xml.XmlNode col in child.ChildNodes)
                {
                    if (col.LocalName != "tableColumn") continue;
                    if (i == index) return col;
                    i++;
                }
            }
            return null;
        }

        private static PdfCellStyle GetFontStyle(ExcelRangeBase cell, PdfCellStyle cellStyle, Dictionary<ExcelTable, ExcelTableNamedStyle> tableStyleCache)
        {
            var cf = cell.ConditionalFormatting.GetConditionalFormattings();
            if (cf != null && cf.Count > 0)
            {
                // Sort ascending — priority 1 beats priority 2, etc.
                var ordered = cf.OrderBy(r => r.Priority);
                foreach (var rule in ordered)
                {
                    if (!rule.ShouldApplyToCell(cell))
                    {
                        if (rule.StopIfTrue) break;
                        continue;
                    }
                    if (rule.Style?.Font != null && rule.Style.Font.HasValue)
                    {
                        cellStyle.dxfFont = rule.Style.Font;
                        return cellStyle; // CF font wins over the table font
                    }
                    if (rule.StopIfTrue) break;
                }
            }
            var tables = cell.Worksheet.Tables.GetIntersectingRanges(cell);
            if (tables.Count > 0)
            {
                var table = tables[0].Value;
                var range = table.Range;
                ExcelTableNamedStyle tableStyle = GetTableStyle(table, tableStyleCache);
                if (tableStyle != null)
                {
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
                    GetTableRegionOverride(cell, table, out var ovCol, out var ovTbl);
                    cellStyle.dxfFontOverride = (ovCol?.Font != null && ovCol.Font.HasValue) ? ovCol.Font
                                              : (ovTbl?.Font != null && ovTbl.Font.HasValue) ? ovTbl.Font
                                              : null;
                    cellStyle.dxfFont = font;
                }
            }
            return cellStyle;
        }

        private static PdfCellAlignmentData GetContentAlignment(ExcelRangeBase cell)
        {
            var contentAlignment = new PdfCellAlignmentData();
            contentAlignment.HorizontalAlignment = (EPPlus.Export.Pdf.Enums.ExcelHorizontalAlignment)cell.Style.HorizontalAlignment;
            contentAlignment.VerticalAlignment = (EPPlus.Export.Pdf.Enums.ExcelVerticalAlignment)cell.Style.VerticalAlignment;
            contentAlignment.Indent = cell.Style.Indent;
            contentAlignment.WrapText = cell.Style.WrapText;
            contentAlignment.ShrinkToFit = cell.Style.ShrinkToFit;
            contentAlignment.TextRotation = (cell.Style.TextRotation > 90) ? ((cell.Style.TextRotation == 255) ? 0 : 90 - cell.Style.TextRotation) : cell.Style.TextRotation;
            contentAlignment.IsVertical = cell.Style.TextRotation == 255 ? true : false;
            contentAlignment.TextDirection = (EPPlus.Export.Pdf.Enums.ExcelReadingOrder)cell.Style.ReadingOrder;
            return contentAlignment;
        }

        private static List<TextFragment> GetTextFragments(PdfPageSettings pageSettings, PdfDictionaries dictionaries, ExcelRangeBase cell, PdfCellStyle cellStyle)
        {
            bool dxfBold, dxfItalic, dxfStrike, dxfUnderline;
            ExcelUnderLineType dxfUnderLineType;
            System.Drawing.Color? dxfColor;
            ReadDxfFontOverrides(cellStyle, out dxfBold, out dxfItalic, out dxfStrike, out dxfUnderline, out dxfUnderLineType, out dxfColor);
            string forcedText = ResolveErrorText(pageSettings, cell);
            if (cell.IsRichText)
            {
                return GetTextFragmentsFromRichText(pageSettings, dictionaries, cell.RichText, forcedText,
                    dxfBold, dxfItalic, dxfStrike, dxfUnderline, dxfUnderLineType, dxfColor);
            }
            return GetTextFragmentsFromCellStyle(pageSettings, dictionaries, cell, forcedText,
                dxfBold, dxfItalic, dxfStrike, dxfUnderline, dxfUnderLineType, dxfColor);
        }

        private static List<TextFragment> GetTextFragmentsFromRichText(PdfPageSettings pageSettings, PdfDictionaries dictionaries, ExcelRichTextCollection richText, string forcedText,
            bool dxfBold, bool dxfItalic, bool dxfStrike, bool dxfUnderline, ExcelUnderLineType dxfUnderLineType, System.Drawing.Color? dxfColor)
        {
            var textFragments = new List<TextFragment>(richText.Count);
            for (int i = 0; i < richText.Count; i++)
            {
                var rt = richText[i];
                var textFrag = new TextFragment();
                textFrag.Font = new RichTextFormatSimple();
                textFrag.Text = forcedText == null ? rt.Text : forcedText;
                textFrag.Font.Family = rt.FontName;
                textFrag.Font.Size = rt.Size;
                textFrag.RichTextOptions.Bold = rt.Bold || dxfBold;
                textFrag.RichTextOptions.Italic = rt.Italic || dxfItalic;
                textFrag.RichTextOptions.UnderlineType = MapUnderlineType(rt.UnderLineType);
                textFrag.RichTextOptions.StrikeType = (rt.Strike || dxfStrike) ? 2 : 1;
                textFrag.RichTextOptions.SuperScript = rt.VerticalAlign == ExcelVerticalAlignmentFont.Superscript;
                textFrag.RichTextOptions.SubScript = rt.VerticalAlign == ExcelVerticalAlignmentFont.Subscript;
                textFrag.RichTextOptions.FontColor = dxfColor ?? rt.Color;
                textFrag.Font.SubFamily = ComputeFontStyle(textFrag);
                textFragments.Add(textFrag);
                dictionaries.AddFont(pageSettings, textFrag.Font.Family, textFrag.Font.SubFamily, textFrag.Text);
            }
            return textFragments;
        }

        private static List<TextFragment> GetTextFragmentsFromCellStyle(PdfPageSettings pageSettings, PdfDictionaries dictionaries, ExcelRangeBase cell, string forcedText,
            bool dxfBold, bool dxfItalic, bool dxfStrike, bool dxfUnderline, ExcelUnderLineType dxfUnderLineType, System.Drawing.Color? dxfColor)
        {
            var font = cell.Style.Font;
            var textFragments = new List<TextFragment>(1);
            var textFrag = new TextFragment();
            textFrag.Font = new RichTextFormatSimple();
            textFrag.Text = forcedText == null ? cell.Text : forcedText;
            textFrag.Font.Family = font.Name;
            textFrag.Font.Size = font.Size;
            textFrag.RichTextOptions.Bold = font.Bold || dxfBold;
            textFrag.RichTextOptions.Italic = font.Italic || dxfItalic;
            // Cell-style underline; dxf overrides only when the cell itself is not underlined.
            ExcelUnderLineType underLineType;
            if (font.UnderLine)
            {
                underLineType = font.UnderLineType == ExcelUnderLineType.None ? ExcelUnderLineType.Single : font.UnderLineType;
            }
            else if (dxfUnderline)
            {
                underLineType = dxfUnderLineType;
            }
            else
            {
                underLineType = ExcelUnderLineType.None;
            }
            textFrag.RichTextOptions.UnderlineType = MapUnderlineType(underLineType);
            textFrag.RichTextOptions.StrikeType = (font.Strike || dxfStrike) ? 2 : 1;
            textFrag.RichTextOptions.SuperScript = font.VerticalAlign == ExcelVerticalAlignmentFont.Superscript;
            textFrag.RichTextOptions.SubScript = font.VerticalAlign == ExcelVerticalAlignmentFont.Subscript;
            textFrag.RichTextOptions.FontColor = dxfColor ?? font.Color.ToColor();
            textFrag.Font.SubFamily = ComputeFontStyle(textFrag);
            textFragments.Add(textFrag);
            dictionaries.AddFont(pageSettings, textFrag.Font.Family, textFrag.Font.SubFamily, textFrag.Text);
            return textFragments;
        }

        private static void ReadDxfFontOverrides(PdfCellStyle cellStyle, out bool bold, out bool italic, out bool strike, out bool underline, out ExcelUnderLineType underLineType, out System.Drawing.Color? color)
        {
            bold = false; italic = false; strike = false; underline = false;
            underLineType = ExcelUnderLineType.None; color = null;
            if (cellStyle == null) return;

            var elem = cellStyle.dxfFont;           // style element (or CF font)
            var ov = cellStyle.dxfFontOverride;   // region override — wins per property

            var b = ov?.Bold ?? elem?.Bold;
            var i = ov?.Italic ?? elem?.Italic;
            var st = ov?.Strike ?? elem?.Strike;
            var un = ov?.Underline ?? elem?.Underline;
            var cl = (ov?.Color != null && ov.Color.HasValue) ? ov.Color
                   : (elem?.Color != null && elem.Color.HasValue) ? elem.Color
                   : null;

            bold = b ?? false;
            italic = i ?? false;
            strike = st ?? false;
            underline = un != null;
            underLineType = un != null ? (ExcelUnderLineType)un : ExcelUnderLineType.None;
            if (cl != null && cl.HasValue)
            {
                // GetColorAsColor() returns white for an automatic colour; for print it must be black.
                color = cl.Auto == true ? System.Drawing.Color.Black : cl.GetColorAsColor();
            }
        }

        private static string ResolveErrorText(PdfPageSettings pageSettings, ExcelRangeBase cell)
        {
            if (!ExcelErrorValue.IsErrorValue(cell.Text)) return null;
            switch (pageSettings.CellErrors)
            {
                case CellErrors.Blank: return "";
                case CellErrors.Dashed: return "--";
                case CellErrors.NA: return "#N/A";
                case CellErrors.Displayed:
                default: return null;
            }
        }

        private static int MapUnderlineType(ExcelUnderLineType type)
        {
            // 12 = none, 13 = single, 4 = double (matches existing rendering code; accounting not supported).
            if (type == ExcelUnderLineType.Single) return 13;
            if (type == ExcelUnderLineType.Double) return 4;
            return 12;
        }

        private static FontSubFamily ComputeFontStyle(TextFragment textFrag)
        {
            if (textFrag.RichTextOptions.Bold && textFrag.RichTextOptions.Italic) return FontSubFamily.BoldItalic;
            if (textFrag.RichTextOptions.Bold) return FontSubFamily.Bold;
            if (textFrag.RichTextOptions.Italic) return FontSubFamily.Italic;
            return FontSubFamily.Regular;
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

        internal static PdfHeaderFooter GetTextFormats(PdfPageSettings pageSettings, PdfDictionaries dictionaries, ExcelWorksheet ws, ExcelHeaderFooterTextCollection textCollection, HeaderFooterType type, HeaderFooterAlignment alignment, HeaderFooterSection section)
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
                    if (pic?.Image?.ImageBytes != null)
                    {
                        imageBytes = pic.Image.ImageBytes;
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
                textFrag.Font.SubFamily = ComputeFontStyle(textFrag);
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

        /// <summary>
        /// Check if we need to add additional columns to accomodate text that is not wrapped and overlaps other cell.s
        /// </summary>
        /// <param name="ws">The worksheet to check.</param>
        /// <returns>The number of columns to add.</returns>
        internal static int AddColumnsForNonWrappedText(PdfPageSettings pageSettings, ExcelWorksheet ws, PdfWorksheet pdfSheet)
        {
            int columnsToAdd = 0;
            var catalog = new PdfCatalog();
            var lastColumn = ws.Dimension.End.Column;
            ExcelRangeBase lastColumnRange = ws.Cells[1, lastColumn, ws.Dimension.End.Row, lastColumn];
            var cc = catalog.GetCellCollectionFromRange(pageSettings, lastColumnRange);
            double textLength = 0;
            for (int i = cc.FromRow; i < cc.ToRow; i++)
            {
                textLength = cc[i, cc.FromColumn].TotalTextLength > textLength ? cc[i, cc.FromColumn].TotalTextLength : textLength;
            }
            double columnWidth = UnitConversion.ExcelColumnWidthToPoints(ws.Column(ws.Dimension._toCol).Width, pdfSheet.ZeroCharWidth);
            while (textLength > columnWidth)
            {
                columnsToAdd++;
                columnWidth += UnitConversion.ExcelColumnWidthToPoints(ws.Column(ws.Dimension._toCol + columnsToAdd).Width, pdfSheet.ZeroCharWidth);
            }
            return columnsToAdd;
        }

        internal static ExcelDxfBorderItem GetTopBorderItem(ExcelRangeBase cell, ExcelBorderItem xfBorder, ExcelTable table, ExcelTableNamedStyle tableStyle, out int elementOrder)
        {
            var range = table.Range;
            int tableRow = cell._fromRow - range._fromRow;
            int tableCol = cell._fromCol - range._fromCol;
            int ts = table.ShowHeader ? 1 : 0;
            var top = tableRow == 0 ? tableStyle.WholeTable.Style.Border.Top : tableStyle.WholeTable.Style.Border.Horizontal;
            elementOrder = TableEdgeOrder.WholeTable;
            if (table.ShowHeader && tableRow == 0)
            {
                if (tableStyle.HeaderRow.Style.Border.Top.HasValue)
                { top = tableStyle.HeaderRow.Style.Border.Top; elementOrder = TableEdgeOrder.HeaderRow; }
                if (tableCol == 0 && table.ShowFirstColumn && tableStyle.FirstHeaderCell.Style.Border.Top.HasValue)
                { top = tableStyle.FirstHeaderCell.Style.Border.Top; elementOrder = TableEdgeOrder.FirstHeaderCell; }
                if (cell._fromCol == range._toCol && table.ShowLastColumn && tableStyle.LastHeaderCell.Style.Border.Top.HasValue)
                { top = tableStyle.LastHeaderCell.Style.Border.Top; elementOrder = TableEdgeOrder.LastHeaderCell; }
            }
            else if (table.ShowTotal && cell._fromRow == range._toRow)
            {
                if (tableStyle.TotalRow.Style.Border.Top.HasValue)
                { top = tableStyle.TotalRow.Style.Border.Top; elementOrder = TableEdgeOrder.TotalRow; }
                if (tableCol == 0 && table.ShowFirstColumn && tableStyle.FirstTotalCell.Style.Border.Top.HasValue)
                { top = tableStyle.FirstTotalCell.Style.Border.Top; elementOrder = TableEdgeOrder.FirstTotalCell; }
                if (cell._fromCol == range._toCol && table.ShowLastColumn && tableStyle.LastTotalCell.Style.Border.Top.HasValue)
                { top = tableStyle.LastTotalCell.Style.Border.Top; elementOrder = TableEdgeOrder.LastTotalCell; }
            }
            else
            {
                if (table.ShowColumnStripes && (tableCol & 1) == 0)
                {
                    if (cell._fromRow - ts > range._fromRow && cell._fromRow < range._toRow)
                    { top = tableStyle.FirstColumnStripe.Style.Border.Horizontal; elementOrder = TableEdgeOrder.FirstColumnStripe; }
                    else if (cell._fromRow <= range._toRow)
                    { top = null; }
                    else
                    { top = tableStyle.FirstColumnStripe.Style.Border.Top; elementOrder = TableEdgeOrder.FirstColumnStripe; }
                }
                if (table.ShowColumnStripes && (tableCol & 1) != 0)
                {
                    if (cell._fromRow + ts > range._fromRow && cell._fromRow < range._toRow)
                    { top = tableStyle.SecondColumnStripe.Style.Border.Horizontal; elementOrder = TableEdgeOrder.SecondColumnStripe; }
                    else if (cell._fromRow <= range._toRow)
                    { top = null; }
                    else
                    { top = tableStyle.SecondColumnStripe.Style.Border.Top; elementOrder = TableEdgeOrder.SecondColumnStripe; }
                }
                if (table.ShowRowStripes && tableStyle.FirstRowStripe.Style.Border.Top.HasValue && (tableRow & 1) != 0)
                { top = tableStyle.FirstRowStripe.Style.Border.Top; elementOrder = TableEdgeOrder.FirstRowStripe; }
                if (table.ShowRowStripes && tableStyle.SecondRowStripe.Style.Border.Top.HasValue && (tableRow & 1) == 0)
                { top = tableStyle.SecondRowStripe.Style.Border.Top; elementOrder = TableEdgeOrder.SecondRowStripe; }
                if (table.ShowLastColumn && tableStyle.LastColumn.Style.Border.Top.HasValue && cell._fromCol == range._toCol)
                { top = tableStyle.LastColumn.Style.Border.Top; elementOrder = TableEdgeOrder.LastColumn; }
                if (table.ShowFirstColumn && tableStyle.FirstColumn.Style.Border.Top.HasValue && tableCol == range._toCol)
                { top = tableStyle.FirstColumn.Style.Border.Top; elementOrder = TableEdgeOrder.FirstColumn; }
            }
            if (top == null || !top.Style.HasValue || top.Style.Value == ExcelBorderStyle.None) elementOrder = TableEdgeOrder.None;
            return top;
        }

        internal static ExcelDxfBorderItem GetBottomBorderItem(ExcelRangeBase cell, ExcelBorderItem xfBorder, ExcelTable table, ExcelTableNamedStyle tableStyle, out int elementOrder)
        {
            var range = table.Range;
            int tableRow = cell._fromRow - range._fromRow;
            int tableCol = cell._fromCol - range._fromCol;
            var bottom = range._toRow == cell._fromRow ? tableStyle.WholeTable.Style.Border.Bottom : null;
            elementOrder = TableEdgeOrder.WholeTable;
            if (table.ShowHeader && tableRow == 0)
            {
                if (tableStyle.HeaderRow.Style.Border.Bottom.HasValue)
                { bottom = tableStyle.HeaderRow.Style.Border.Bottom; elementOrder = TableEdgeOrder.HeaderRow; }
                if (tableCol == 0 && table.ShowFirstColumn && tableStyle.FirstHeaderCell.Style.Border.Bottom.HasValue)
                { bottom = tableStyle.FirstHeaderCell.Style.Border.Bottom; elementOrder = TableEdgeOrder.FirstHeaderCell; }
                if (cell._fromCol == range._toCol && table.ShowLastColumn && tableStyle.LastHeaderCell.Style.Border.Bottom.HasValue)
                { bottom = tableStyle.LastHeaderCell.Style.Border.Bottom; elementOrder = TableEdgeOrder.LastHeaderCell; }
            }
            else if (table.ShowTotal && cell._fromRow == range._toRow)
            {
                if (tableStyle.TotalRow.Style.Border.Bottom.HasValue)
                { bottom = tableStyle.TotalRow.Style.Border.Bottom; elementOrder = TableEdgeOrder.TotalRow; }
                if (tableCol == 0 && table.ShowFirstColumn && tableStyle.FirstTotalCell.Style.Border.Bottom.HasValue)
                { bottom = tableStyle.FirstTotalCell.Style.Border.Bottom; elementOrder = TableEdgeOrder.FirstTotalCell; }
                if (cell._fromCol == range._toCol && table.ShowLastColumn && tableStyle.LastTotalCell.Style.Border.Bottom.HasValue)
                { bottom = tableStyle.LastTotalCell.Style.Border.Bottom; elementOrder = TableEdgeOrder.LastTotalCell; }
            }
            else
            {
                if (table.ShowColumnStripes && tableStyle.FirstColumnStripe.Style.Border.Bottom.HasValue && (tableCol & 1) != 0)
                {
                    if (cell._fromRow > range._fromRow && cell._fromRow < range._toRow)
                    { bottom = tableStyle.FirstColumnStripe.Style.Border.Horizontal; elementOrder = TableEdgeOrder.FirstColumnStripe; }
                    else if (cell._fromRow < range._toRow)
                    { bottom = null; }
                    else
                    { bottom = tableStyle.FirstColumnStripe.Style.Border.Bottom; elementOrder = TableEdgeOrder.FirstColumnStripe; }
                }
                if (table.ShowColumnStripes && tableStyle.SecondColumnStripe.Style.Border.Bottom.HasValue && (tableCol & 1) == 0)
                {
                    if (cell._fromRow > range._fromRow && cell._fromRow < range._toRow)
                    { bottom = tableStyle.SecondColumnStripe.Style.Border.Horizontal; elementOrder = TableEdgeOrder.SecondColumnStripe; }
                    else if (cell._fromRow < range._toRow)
                    { bottom = null; }
                    else
                    { bottom = tableStyle.SecondColumnStripe.Style.Border.Bottom; elementOrder = TableEdgeOrder.SecondColumnStripe; }
                }
                if (table.ShowRowStripes && tableStyle.FirstRowStripe.Style.Border.Bottom.HasValue && (tableRow & 1) != 0)
                { bottom = tableStyle.FirstRowStripe.Style.Border.Bottom; elementOrder = TableEdgeOrder.FirstRowStripe; }
                if (table.ShowRowStripes && tableStyle.SecondRowStripe.Style.Border.Bottom.HasValue && (tableRow & 1) == 0)
                { bottom = tableStyle.SecondRowStripe.Style.Border.Bottom; elementOrder = TableEdgeOrder.SecondRowStripe; }
                if (table.ShowLastColumn && tableStyle.LastColumn.Style.Border.Bottom.HasValue && cell._fromCol == range._toCol)
                { bottom = tableStyle.LastColumn.Style.Border.Bottom; elementOrder = TableEdgeOrder.LastColumn; }
                if (table.ShowFirstColumn && tableStyle.FirstColumn.Style.Border.Bottom.HasValue && tableCol == 0)
                { bottom = tableStyle.FirstColumn.Style.Border.Bottom; elementOrder = TableEdgeOrder.FirstColumn; }
            }
            if (bottom == null || !bottom.Style.HasValue || bottom.Style.Value == ExcelBorderStyle.None) elementOrder = TableEdgeOrder.None;
            return bottom;
        }

        internal static ExcelDxfBorderItem GetLeftBorderItem(ExcelRangeBase cell, ExcelBorderItem xfBorder, ExcelTable table, ExcelTableNamedStyle tableStyle, out int elementOrder)
        {
            var range = table.Range;
            int tableRow = cell._fromRow - range._fromRow;
            int tableCol = cell._fromCol - range._fromCol;
            var left = tableCol == 0 ? tableStyle.WholeTable.Style.Border.Left : tableStyle.WholeTable.Style.Border.Vertical;
            elementOrder = TableEdgeOrder.WholeTable;
            if (table.ShowHeader && tableRow == 0)
            {
                if (tableStyle.HeaderRow.Style.Border.Left.HasValue)
                { left = tableStyle.HeaderRow.Style.Border.Left; elementOrder = TableEdgeOrder.HeaderRow; }
                if (tableCol == 0 && table.ShowFirstColumn && tableStyle.FirstHeaderCell.Style.Border.Left.HasValue)
                { left = tableStyle.FirstHeaderCell.Style.Border.Left; elementOrder = TableEdgeOrder.FirstHeaderCell; }
                if (cell._fromCol == range._toCol && table.ShowLastColumn && tableStyle.LastHeaderCell.Style.Border.Left.HasValue)
                { left = tableStyle.LastHeaderCell.Style.Border.Left; elementOrder = TableEdgeOrder.LastHeaderCell; }
            }
            else if (table.ShowTotal && cell._fromRow == range._toRow)
            {
                if (tableStyle.TotalRow.Style.Border.Left.HasValue)
                { left = tableStyle.TotalRow.Style.Border.Left; elementOrder = TableEdgeOrder.TotalRow; }
                if (tableCol == 0 && table.ShowFirstColumn && tableStyle.FirstTotalCell.Style.Border.Left.HasValue)
                { left = tableStyle.FirstTotalCell.Style.Border.Left; elementOrder = TableEdgeOrder.FirstTotalCell; }
                if (cell._fromCol == range._toCol && table.ShowLastColumn && tableStyle.LastTotalCell.Style.Border.Left.HasValue)
                { left = tableStyle.LastTotalCell.Style.Border.Left; elementOrder = TableEdgeOrder.LastTotalCell; }
            }
            else
            {
                if (table.ShowColumnStripes && tableStyle.FirstColumnStripe.Style.Border.Left.HasValue && (tableCol & 1) != 0)
                { left = tableStyle.FirstColumnStripe.Style.Border.Left; elementOrder = TableEdgeOrder.FirstColumnStripe; }
                if (table.ShowColumnStripes && tableStyle.SecondColumnStripe.Style.Border.Left.HasValue && (tableCol & 1) == 0)
                { left = tableStyle.SecondColumnStripe.Style.Border.Left; elementOrder = TableEdgeOrder.SecondColumnStripe; }
                if (table.ShowRowStripes && tableStyle.FirstRowStripe.Style.Border.Left.HasValue && (tableRow & 1) != 0)
                {
                    if (cell._fromCol > range._fromCol)
                    {
                        left = tableStyle.FirstRowStripe.Style.Border.Vertical;
                        elementOrder = TableEdgeOrder.FirstRowStripe;
                    }
                    else
                    {
                        left = tableStyle.FirstRowStripe.Style.Border.Left;
                        elementOrder = TableEdgeOrder.FirstRowStripe;
                    }
                }
                if (table.ShowRowStripes && tableStyle.SecondRowStripe.Style.Border.Left.HasValue && (tableRow & 1) == 0)
                {
                    if (cell._fromCol > range._fromCol)
                    {
                        left = tableStyle.SecondRowStripe.Style.Border.Vertical;
                        elementOrder = TableEdgeOrder.SecondRowStripe;
                    }
                    else
                    {
                        left = tableStyle.SecondRowStripe.Style.Border.Left;
                        elementOrder = TableEdgeOrder.SecondRowStripe;
                    }
                }
                if (table.ShowLastColumn && tableStyle.LastColumn.Style.Border.Left.HasValue && cell._fromCol == range._toCol)
                { left = tableStyle.LastColumn.Style.Border.Left; elementOrder = TableEdgeOrder.LastColumn; }
                if (table.ShowFirstColumn && tableStyle.FirstColumn.Style.Border.Left.HasValue && tableCol == range._toCol)
                { left = tableStyle.FirstColumn.Style.Border.Left; elementOrder = TableEdgeOrder.FirstColumn; }
            }
            if (left == null || !left.Style.HasValue || left.Style.Value == ExcelBorderStyle.None) elementOrder = TableEdgeOrder.None;
            return left;
        }

        internal static ExcelDxfBorderItem GetRightBorderItem(ExcelRangeBase cell, ExcelBorderItem xfBorder, ExcelTable table, ExcelTableNamedStyle tableStyle, out int elementOrder)
        {
            var range = table.Range;
            int tableRow = cell._fromRow - range._fromRow;
            int tableCol = cell._fromCol - range._fromCol;
            var right = cell._fromCol == range._toCol ? tableStyle.WholeTable.Style.Border.Right : tableStyle.WholeTable.Style.Border.Vertical;
            elementOrder = TableEdgeOrder.WholeTable;
            if (table.ShowHeader && tableRow == 0)
            {
                if (tableStyle.HeaderRow.Style.Border.Right.HasValue)
                { right = tableStyle.HeaderRow.Style.Border.Right; elementOrder = TableEdgeOrder.HeaderRow; }
                if (tableCol == 0 && table.ShowFirstColumn && tableStyle.FirstHeaderCell.Style.Border.Right.HasValue)
                { right = tableStyle.FirstHeaderCell.Style.Border.Right; elementOrder = TableEdgeOrder.FirstHeaderCell; }
                if (cell._fromCol == range._toCol && table.ShowLastColumn && tableStyle.LastHeaderCell.Style.Border.Right.HasValue)
                { right = tableStyle.LastHeaderCell.Style.Border.Right; elementOrder = TableEdgeOrder.LastHeaderCell; }
            }
            else if (table.ShowTotal && cell._fromRow == range._toRow)
            {
                if (tableStyle.TotalRow.Style.Border.Right.HasValue)
                { right = tableStyle.TotalRow.Style.Border.Right; elementOrder = TableEdgeOrder.TotalRow; }
                if (tableCol == 0 && table.ShowFirstColumn && tableStyle.FirstTotalCell.Style.Border.Right.HasValue)
                { right = tableStyle.FirstTotalCell.Style.Border.Right; elementOrder = TableEdgeOrder.FirstTotalCell; }
                if (cell._fromCol == range._toCol && table.ShowLastColumn && tableStyle.LastTotalCell.Style.Border.Right.HasValue)
                { right = tableStyle.LastTotalCell.Style.Border.Right; elementOrder = TableEdgeOrder.LastTotalCell; }
            }
            else
            {
                if (table.ShowColumnStripes && tableStyle.FirstColumnStripe.Style.Border.Right.HasValue && (tableCol & 1) != 0)
                { right = tableStyle.FirstColumnStripe.Style.Border.Right; elementOrder = TableEdgeOrder.FirstColumnStripe; }
                if (table.ShowColumnStripes && tableStyle.SecondColumnStripe.Style.Border.Right.HasValue && (tableCol & 1) == 0)
                { right = tableStyle.SecondColumnStripe.Style.Border.Right; elementOrder = TableEdgeOrder.SecondColumnStripe; }
                if (table.ShowRowStripes && tableStyle.FirstRowStripe.Style.Border.Right.HasValue && (tableRow & 1) != 0)
                {
                    if (cell._fromCol < range._toCol)
                    {
                        right = tableStyle.FirstRowStripe.Style.Border.Vertical;
                        elementOrder = TableEdgeOrder.FirstRowStripe;
                    }
                    else
                    {
                        right = tableStyle.FirstRowStripe.Style.Border.Right;
                        elementOrder = TableEdgeOrder.FirstRowStripe;
                    }
                }
                if (table.ShowRowStripes && tableStyle.SecondRowStripe.Style.Border.Right.HasValue && (tableRow & 1) == 0)
                {
                    if (cell._fromCol < range._toCol)
                    {
                        right = tableStyle.FirstRowStripe.Style.Border.Vertical;
                        elementOrder = TableEdgeOrder.FirstRowStripe;
                    }
                    else
                    {
                        right = tableStyle.FirstRowStripe.Style.Border.Right;
                        elementOrder = TableEdgeOrder.FirstRowStripe;
                    }
                }
                if (table.ShowLastColumn && tableStyle.LastColumn.Style.Border.Right.HasValue && cell._fromCol == range._toCol)
                { right = tableStyle.LastColumn.Style.Border.Right; elementOrder = TableEdgeOrder.LastColumn; }
                if (table.ShowFirstColumn && tableStyle.FirstColumn.Style.Border.Right.HasValue && tableCol == range._toCol)
                { right = tableStyle.FirstColumn.Style.Border.Right; elementOrder = TableEdgeOrder.FirstColumn; }
            }
            if (right == null || !right.Style.HasValue || right.Style.Value == ExcelBorderStyle.None) elementOrder = TableEdgeOrder.None;
            return right;
        }

        private static void ReconcileSharedBorders(PdfCellCollection map)
        {
            for (int row = map.FromRow; row <= map.ToRow; row++)
            {
                for (int col = map.FromColumn; col <= map.ToColumn; col++)
                {
                    var cell = map[row, col];
                    if (cell == null || cell.Hidden || cell.Merged || cell.CellStyle == null) continue;
                    var cs = cell.CellStyle;

                    // Vertical shared edge: this cell's right vs the next cell's left.
                    if (col < map.ToColumn)
                    {
                        var next = map[row, col + 1];
                        if (next != null && !next.Hidden && !next.Merged && next.CellStyle != null)
                        {
                            var ns = next.CellStyle;
                            // Two adjacent DOUBLE borders form ONE shared double: keep BOTH sides so each
                            // cell draws only its inner line (see PdfBorderRenderer.DrawDoubleBorder /
                            // NeighborDouble). Suppressing either side collapses it to a single line.
                            if (!(IsDoubleEdge(cs.xfRight, cs.dxfRight) && IsDoubleEdge(ns.xfLeft, ns.dxfLeft)))
                            {
                                int here = EdgeRank(cs.xfRight, cs.dxfRight, cs.dxfRightElementOrder);
                                int there = EdgeRank(ns.xfLeft, ns.dxfLeft, ns.dxfLeftElementOrder);
                                if (here >= there) ns.SuppressLeft = true;    // this cell's right wins
                                else cs.SuppressRight = true;   // neighbour's left wins
                            }
                        }
                    }

                    // Horizontal shared edge: this cell's bottom vs the cell below's top.
                    if (row < map.ToRow)
                    {
                        var below = map[row + 1, col];
                        if (below != null && !below.Hidden && !below.Merged && below.CellStyle != null)
                        {
                            var bs = below.CellStyle;
                            // Two adjacent DOUBLE borders form ONE shared double: keep BOTH sides (inner-only each).
                            if (!(IsDoubleEdge(cs.xfBottom, cs.dxfBottom) && IsDoubleEdge(bs.xfTop, bs.dxfTop)))
                            {
                                int here = EdgeRank(cs.xfBottom, cs.dxfBottom, cs.dxfBottomElementOrder);
                                int there = EdgeRank(bs.xfTop, bs.dxfTop, bs.dxfTopElementOrder);
                                if (here >= there) bs.SuppressTop = true;     // this cell's bottom wins
                                else cs.SuppressBottom = true;  // cell-below's top wins
                            }
                        }
                    }
                }
            }
        }

        // Effective style of one edge is Double (user xf wins over conditional dxf). Mirrors PdfLayout.IsDouble*.
        private static bool IsDoubleEdge(ExcelBorderItem xf, ExcelDxfBorderItem dxf)
        {
            if (xf != null && xf.Style != ExcelBorderStyle.None) return xf.Style == ExcelBorderStyle.Double;
            if (dxf != null && dxf.Style.HasValue) return dxf.Style.Value == ExcelBorderStyle.Double;
            return false;
        }

        private static int EdgeRank(ExcelBorderItem xf, ExcelDxfBorderItem dxf, int elementOrder)
        {
            // User-applied (xf) border is the highest source.
            if (xf != null && xf.Style != ExcelBorderStyle.None)
                return TableEdgeOrder.UserSet;
            // Otherwise rank purely by where it came from (CF or table element order).
            if (dxf != null && dxf.Style.HasValue && dxf.Style.Value != ExcelBorderStyle.None)
                return elementOrder;
            return 0; // no border
        }

        internal static class TableEdgeOrder
        {
            public const int None = 0, WholeTable = 1,
                FirstColumnStripe = 2, SecondColumnStripe = 3,
                FirstRowStripe = 4, SecondRowStripe = 5,
                LastColumn = 6, FirstColumn = 7,
                TotalRow = 8, HeaderRow = 9,
                FirstHeaderCell = 10, LastHeaderCell = 11,
                FirstTotalCell = 12, LastTotalCell = 13,
                ConditionalFormat = 50,   // beats any table element
                UserSet = 100;            // beats CF and table
        }

        private static ExcelTableNamedStyle GetTableStyle(ExcelTable table, Dictionary<ExcelTable, ExcelTableNamedStyle> cache)
        {
            if (cache.TryGetValue(table, out var cached))
                return cached;

            ExcelTableNamedStyle tableStyle;
            if (table.TableStyle == TableStyles.Custom)
            {
                tableStyle = table.WorkSheet.Workbook.Styles.TableStyles[table.StyleName].As.TableStyle;
            }
            else
            {
                var tmpNode = table.WorkSheet.Workbook.StylesXml.CreateElement("c:tableStyle");
                tableStyle = new ExcelTableNamedStyle(
                    table.WorkSheet.Workbook.Styles.NameSpaceManager, tmpNode, table.WorkSheet.Workbook.Styles);
                tableStyle.SetFromTemplate((TableStyles)table.TableStyle);
            }

            cache[table] = tableStyle;
            return tableStyle;
        }

        private static bool FillIsPaintable(ExcelDxfFill fill)
        {
            if (fill == null || !fill.HasValue)
                return false;

            // SetFill treats a null PatternType as Solid.
            var pattern = fill.PatternType != null
                ? (ExcelFillStyle)fill.PatternType
                : ExcelFillStyle.Solid;

            if (pattern == ExcelFillStyle.None)
                return fill.Gradient != null;                               // only a gradient paints when pattern is None

            if (pattern == ExcelFillStyle.Solid)
                return !string.IsNullOrEmpty(fill.BackgroundColor?.LookupColor());  // Solid needs a real colour

            return true;                                                    // any other pattern type paints
        }
    }
}
