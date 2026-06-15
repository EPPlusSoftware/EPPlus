using EPPlus.Export.Pdf.PdfLayout;
using EPPlus.Export.Pdf.PdfResources;
using EPPlus.Export.Pdf.PdfSettings;
using EPPlus.Fonts.OpenType;
using EPPlus.Fonts.OpenType.Integration;
using EPPlus.Fonts.OpenType.Integration.DataHolders;
using EPPlus.Graphics.Units;
using OfficeOpenXml;
using OfficeOpenXml.ConditionalFormatting;
using OfficeOpenXml.ConditionalFormatting.Contracts;
using OfficeOpenXml.Interfaces.Drawing.Text;
using OfficeOpenXml.Interfaces.Fonts;
using OfficeOpenXml.Interfaces.RichText;
using OfficeOpenXml.Style;
using OfficeOpenXml.Style.HeaderFooterTextFormat;
using OfficeOpenXml.Style.Table;
using OfficeOpenXml.Table;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;

namespace EPPlus.Export.Pdf.PdfCatalog
{
    internal class PdfTextMap
    {
        public static PdfCellCollection SetTextMap(PdfPageSettings pageSettings, PdfDictionaries dictionaries, PdfWorksheet pdfSheet, ref PdfRange pdfRange)
        {
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
                        HandleMergedCell(pageSettings, dictionaries, cell, checkedMergedCells, Map, tempMap, pdfSheet.ZeroCharWidth);
                    }
                    var cellStyle = new PdfCellStyle();
                    GetBorderStyles(cell, cellStyle, tempMap);
                    if (tempMap.Main == null)
                    {
                        GetFillStyles(cell, cellStyle);
                        GetFontStyle(cell, cellStyle);
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
            pdfRange = Range;
            return Map;
        }

        private static void HandleMergedCell(PdfPageSettings pageSettings, PdfDictionaries dictionaries, ExcelRange cell, List<string> checkedMergedCells, PdfCellCollection map, PdfCell tempMap, double ZeroCharWidth)
        {
            var worksheet = cell.Worksheet;
            string mergeAddress = worksheet.MergedCells[cell.Start.Row, cell.Start.Column];
            ExcelAddressBase address = new ExcelAddressBase(mergeAddress);
            if (!checkedMergedCells.Contains(mergeAddress))
            {
                double totalWidth = 0, totalHeight = 0;
                for (int k = address._fromRow; k <= address._toRow; k++)
                {
                    totalHeight += UnitConversion.ExcelRowHeightToPoints(worksheet.Row(k).Height);
                }
                for (int l = address._fromCol; l <= address._toCol; l++)
                {
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
                    GetBorderStyles(main, cellStyle, mainCell);
                    GetFillStyles(main, cellStyle);
                    GetFontStyle(main, cellStyle);

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

        private static void GetFillStyles(ExcelRangeBase cell, PdfCellStyle cellStyle)
        {
            if (cell.Style.Fill.IsEmpty())
            {
                //Conditional Formating
                var cf = cell.ConditionalFormatting.GetConditionalFormattings();
                if (cf != null && cf.Count > 1)
                {
                    // Sort ascending — priority 1 beats priority 2, etc.
                    var ordered = cf.OrderBy(r => r.Priority);

                    foreach (var rule in ordered)
                    {
                        if (!EvaluateConditionalFormattingRule(rule, cell))
                        {
                            if (rule.StopIfTrue) break; // higher-priority rule fired but had no fill — stop anyway
                            continue;
                        }

                        if (rule.Style?.Fill != null && rule.Style.Fill.HasValue)
                        {
                            cellStyle.dxfFill = rule.Style.Fill;
                            return; // CF fill wins — skip table and xf entirely
                        }

                        if (rule.StopIfTrue) break; // rule applied but defined no fill — stop evaluating lower rules
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
                    tableRow = cell._fromRow - range._fromRow;
                    tableCol = cell._fromCol - range._fromCol;
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
                        cellStyle.dxfFill = (tableRow & 1) == 0 ? tableStyle.SecondRowStripe.Style.Fill : tableStyle.FirstRowStripe.Style.Fill;
                    }
                    else if (table.ShowColumnStripes)
                    {
                        cellStyle.dxfFill = (tableCol & 1) != 0 ? tableStyle.SecondColumnStripe.Style.Fill : tableStyle.FirstColumnStripe.Style.Fill;
                    }
                    else
                    {
                        cellStyle.dxfFill = tableStyle.WholeTable.Style.Fill;
                    }
                }
            }
            cellStyle.xfFill = cell.Style.Fill;
        }
        private static bool EvaluateConditionalFormattingRule(IExcelConditionalFormattingRule rule, ExcelRangeBase cell)
        {
            // These never produce a dxf fill — skip entirely
            switch (rule.Type)
            {
                case eExcelConditionalFormattingRuleType.TwoColorScale:
                case eExcelConditionalFormattingRuleType.ThreeColorScale:
                case eExcelConditionalFormattingRuleType.DataBar:
                case eExcelConditionalFormattingRuleType.ThreeIconSet:
                case eExcelConditionalFormattingRuleType.FourIconSet:
                case eExcelConditionalFormattingRuleType.FiveIconSet:
                    return false;
            }

            var cellValue = cell.Value;
            var cellText = cellValue?.ToString() ?? string.Empty;

            switch (rule.Type)
            {
                // --- Comparison (CellIs) types ---
                case eExcelConditionalFormattingRuleType.Equal:
                case eExcelConditionalFormattingRuleType.NotEqual:
                case eExcelConditionalFormattingRuleType.GreaterThan:
                case eExcelConditionalFormattingRuleType.GreaterThanOrEqual:
                case eExcelConditionalFormattingRuleType.LessThan:
                case eExcelConditionalFormattingRuleType.LessThanOrEqual:
                case eExcelConditionalFormattingRuleType.Between:
                case eExcelConditionalFormattingRuleType.NotBetween:
                    return EvaluateCellIs(rule, cell);

                // --- Text types ---
                case eExcelConditionalFormattingRuleType.ContainsText:
                    return cellText.Contains((rule as IExcelConditionalFormattingContainsText)?.Text ?? "");

                case eExcelConditionalFormattingRuleType.NotContainsText:
                    return !cellText.Contains((rule as IExcelConditionalFormattingNotContainsText)?.Text ?? "");

                case eExcelConditionalFormattingRuleType.NotContains:
                    return !cellText.Contains((rule as IExcelConditionalFormattingContainsText)?.Text ?? "");

                case eExcelConditionalFormattingRuleType.BeginsWith:
                    return cellText.StartsWith((rule as IExcelConditionalFormattingBeginsWith)?.Text ?? "");

                case eExcelConditionalFormattingRuleType.EndsWith:
                    return cellText.EndsWith((rule as IExcelConditionalFormattingEndsWith)?.Text ?? "");

                // --- Blank/Error types ---
                case eExcelConditionalFormattingRuleType.ContainsBlanks:
                    return string.IsNullOrEmpty(cellText.Trim());

                case eExcelConditionalFormattingRuleType.NotContainsBlanks:
                    return !string.IsNullOrEmpty(cellText.Trim());

                case eExcelConditionalFormattingRuleType.ContainsErrors:
                    return cellValue is ExcelErrorValue;

                case eExcelConditionalFormattingRuleType.NotContainsErrors:
                    return !(cellValue is ExcelErrorValue);

                // --- Duplicate/Unique — require range-wide scan, stub for now ---
                case eExcelConditionalFormattingRuleType.DuplicateValues:
                case eExcelConditionalFormattingRuleType.UniqueValues:
                    return false;

                // --- Average/StdDev/Top/Bottom/TimePeriod — require range-wide scan, stub for now ---
                case eExcelConditionalFormattingRuleType.AboveAverage:
                case eExcelConditionalFormattingRuleType.AboveOrEqualAverage:
                case eExcelConditionalFormattingRuleType.BelowAverage:
                case eExcelConditionalFormattingRuleType.BelowOrEqualAverage:
                case eExcelConditionalFormattingRuleType.AboveStdDev:
                case eExcelConditionalFormattingRuleType.BelowStdDev:
                case eExcelConditionalFormattingRuleType.Top:
                case eExcelConditionalFormattingRuleType.TopPercent:
                case eExcelConditionalFormattingRuleType.Bottom:
                case eExcelConditionalFormattingRuleType.BottomPercent:
                case eExcelConditionalFormattingRuleType.Last7Days:
                case eExcelConditionalFormattingRuleType.LastMonth:
                case eExcelConditionalFormattingRuleType.LastWeek:
                case eExcelConditionalFormattingRuleType.NextMonth:
                case eExcelConditionalFormattingRuleType.NextWeek:
                case eExcelConditionalFormattingRuleType.ThisMonth:
                case eExcelConditionalFormattingRuleType.ThisWeek:
                case eExcelConditionalFormattingRuleType.Today:
                case eExcelConditionalFormattingRuleType.Tomorrow:
                case eExcelConditionalFormattingRuleType.Yesterday:
                    return false;

                // --- Formula-based ---
                case eExcelConditionalFormattingRuleType.Expression:
                    return false; // requires formula evaluation — future phase

                default:
                    return false;
            }
        }

        private static bool EvaluateCellIs(IExcelConditionalFormattingRule rule, ExcelRangeBase cell)
        {
            if (cell.Value == null) return false;
            if (!double.TryParse(cell.Value.ToString(), out double cellDouble)) return false;

            var f1Rule = rule as IExcelConditionalFormattingWithFormula;
            if (f1Rule == null || !double.TryParse(f1Rule.Formula, out double formula1)) return false;

            switch (rule.Type)
            {
                case eExcelConditionalFormattingRuleType.Equal: return cellDouble == formula1;
                case eExcelConditionalFormattingRuleType.NotEqual: return cellDouble != formula1;
                case eExcelConditionalFormattingRuleType.GreaterThan: return cellDouble > formula1;
                case eExcelConditionalFormattingRuleType.GreaterThanOrEqual: return cellDouble >= formula1;
                case eExcelConditionalFormattingRuleType.LessThan: return cellDouble < formula1;
                case eExcelConditionalFormattingRuleType.LessThanOrEqual: return cellDouble <= formula1;
                case eExcelConditionalFormattingRuleType.Between:
                case eExcelConditionalFormattingRuleType.NotBetween:
                    {
                        var f2Rule = rule as IExcelConditionalFormattingWithFormula2;
                        if (f2Rule == null || !double.TryParse(f2Rule.Formula2, out double formula2)) return false;
                        bool inRange = cellDouble >= formula1 && cellDouble <= formula2;
                        return rule.Type == eExcelConditionalFormattingRuleType.Between ? inRange : !inRange;
                    }
                default:
                    return false;
            }
        }

        private static void GetBorderStyles(ExcelRangeBase cell, PdfCellStyle cellStyle, PdfCell pcell)
        {

            /* Kika på varje del av border top bottom left right
             * om cell har top använd den
             * om cell inte har border, gå igenom prio ordning på tabell borders och använd WholeTable om null.
             * om cellein i tabellen är i mitten av tabellen använd horizontal och vertical border istället.
             * I fallet top så ska om vi är i header row eller om cell fromrow är samma som table fromrow så ska top border vara top border. annar är det horizontal som gäller
             * Glöm ej vertical border om fromcol är samma som tabell fromcol.
             */
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
                    cellStyle.dxfTop = PdfCellBorderLayout.GetTopBorderItem(cell, cellStyle.xfTop, table, tableStyle);
                    cellStyle.dxfBottom = PdfCellBorderLayout.GetBottomBorderItem(cell, cellStyle.xfBottom, table, tableStyle);
                    cellStyle.dxfLeft = PdfCellBorderLayout.GetLeftBorderItem(cell, cellStyle.xfLeft, table, tableStyle);
                    cellStyle.dxfRight = PdfCellBorderLayout.GetRightBorderItem(cell, cellStyle.xfRight, table, tableStyle);
                }
            }
        }

        private static PdfCellStyle GetFontStyle(ExcelRangeBase cell, PdfCellStyle cellStyle)
        {
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
            contentAlignment.TextRotation = (cell.Style.TextRotation > 90) ? ((cell.Style.TextRotation == 255) ? 0 : 90 - cell.Style.TextRotation) : cell.Style.TextRotation;
            contentAlignment.IsVertical = cell.Style.TextRotation == 255 ? true : false;
            contentAlignment.TextDirection = cell.Style.ReadingOrder;
            return contentAlignment;
        }

        internal static long subFamilyTicks;
        internal static long addFontTicks;
        internal static int subFamilyCount;

        private static List<TextFragment> GetTextFragments(PdfPageSettings pageSettings, PdfDictionaries dictionaries, ExcelRangeBase cell, PdfCellStyle cellStyle)
        {
            bool dxfBold, dxfItalic, dxfStrike, dxfUnderline;
            ExcelUnderLineType dxfUnderLineType;
            ReadDxfFontOverrides(cellStyle, out dxfBold, out dxfItalic, out dxfStrike, out dxfUnderline, out dxfUnderLineType);

            string forcedText = ResolveErrorText(pageSettings, cell);

            if (cell.IsRichText)
            {
                return GetTextFragmentsFromRichText(pageSettings, dictionaries, cell.RichText, forcedText,
                    dxfBold, dxfItalic, dxfStrike, dxfUnderline, dxfUnderLineType);
            }
            return GetTextFragmentsFromCellStyle(pageSettings, dictionaries, cell, forcedText,
                dxfBold, dxfItalic, dxfStrike, dxfUnderline, dxfUnderLineType);
        }

        private static List<TextFragment> GetTextFragmentsFromRichText(PdfPageSettings pageSettings, PdfDictionaries dictionaries, ExcelRichTextCollection richText, string forcedText,
            bool dxfBold, bool dxfItalic, bool dxfStrike, bool dxfUnderline, ExcelUnderLineType dxfUnderLineType)
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
                textFrag.RichTextOptions.FontColor = rt.Color;

                textFrag.Font.SubFamily = ComputeFontStyle(textFrag);

                textFragments.Add(textFrag);
                dictionaries.AddFont(pageSettings, textFrag.Font.Family, textFrag.Font.SubFamily, textFrag.Text);
            }
            return textFragments;
        }

        private static List<TextFragment> GetTextFragmentsFromCellStyle(PdfPageSettings pageSettings, PdfDictionaries dictionaries, ExcelRangeBase cell, string forcedText,
            bool dxfBold, bool dxfItalic, bool dxfStrike, bool dxfUnderline, ExcelUnderLineType dxfUnderLineType)
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
            textFrag.RichTextOptions.FontColor = font.Color.ToColor();

            textFrag.Font.SubFamily = ComputeFontStyle(textFrag);

            textFragments.Add(textFrag);
            dictionaries.AddFont(pageSettings, textFrag.Font.Family, textFrag.Font.SubFamily, textFrag.Text);
            return textFragments;
        }

        private static void ReadDxfFontOverrides(PdfCellStyle cellStyle, out bool bold, out bool italic, out bool strike, out bool underline, out ExcelUnderLineType underLineType)
        {
            bold = false;
            italic = false;
            strike = false;
            underline = false;
            underLineType = ExcelUnderLineType.None;
            if (cellStyle != null && cellStyle.dxfFont != null)
            {
                bold = cellStyle.dxfFont.Bold != null ? (bool)cellStyle.dxfFont.Bold : false;
                italic = cellStyle.dxfFont.Italic != null ? (bool)cellStyle.dxfFont.Italic : false;
                strike = cellStyle.dxfFont.Strike != null ? (bool)cellStyle.dxfFont.Strike : false;
                underline = cellStyle.dxfFont.Underline != null;
                underLineType = cellStyle.dxfFont.Underline != null ? (ExcelUnderLineType)cellStyle.dxfFont.Underline : ExcelUnderLineType.None;
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

        //private static MeasurementFontStyles ComputeFontStyle(TextFragment textFrag)
        //{
        //    return (textFrag.RichTextOptions.Bold ? MeasurementFontStyles.Bold : 0) |
        //           (textFrag.RichTextOptions.Italic ? MeasurementFontStyles.Italic : 0) |
        //           (textFrag.RichTextOptions.UnderlineType != 12 ? MeasurementFontStyles.Underline : 0) |
        //           (textFrag.RichTextOptions.StrikeType > 1 ? MeasurementFontStyles.Strikeout : 0);
        //}

        private static FontSubFamily ComputeFontStyle(TextFragment textFrag)
        {
            if (textFrag.RichTextOptions.Bold && textFrag.RichTextOptions.Italic) return FontSubFamily.BoldItalic;
            if(textFrag.RichTextOptions.Bold) return FontSubFamily.Bold;
            if (textFrag.RichTextOptions.Italic) return FontSubFamily.Italic;
            return FontSubFamily.Regular;
        }


        //private static List<PdfTextFormat> GetTextFormats(PdfPageSettings pageSettings, PdfDictionaries dictionaries, ExcelRichTextCollection RichTextCollection, PdfCellStyle cellStyle = null)
        //{
        //    var textFormats = new List<PdfTextFormat>();
        //    bool bold = false, italic = false, underline = false, strike = false;
        //    ExcelUnderLineType underLineType = ExcelUnderLineType.None;
        //    if (cellStyle != null && cellStyle.dxfFont != null)
        //    {
        //        bold = cellStyle.dxfFont.Bold != null ? (bool)cellStyle.dxfFont.Bold : false;
        //        italic = cellStyle.dxfFont.Italic != null ? (bool)cellStyle.dxfFont.Italic : false;
        //        strike = cellStyle.dxfFont.Strike != null ? (bool)cellStyle.dxfFont.Strike : false;
        //        underline = cellStyle.dxfFont.Underline != null;
        //        underLineType = cellStyle.dxfFont.Underline != null ? (ExcelUnderLineType)cellStyle.dxfFont.Underline : ExcelUnderLineType.None;
        //    }
        //    for (int i = 0; i < RichTextCollection.Count; i++)
        //    {
        //        var rt = RichTextCollection[i];
        //        var textFormat = new PdfTextFormat();
        //        textFormat.Text = rt.Text;
        //        textFormat.FontName = rt.FontName;
        //        textFormat.FontFamily = rt.Family;
        //        textFormat.FontSize = rt.Size;
        //        textFormat.Bold = rt.Bold || bold;
        //        textFormat.Italic = rt.Italic || italic;
        //        textFormat.Strike = rt.Strike || strike;
        //        textFormat.Underline = rt.UnderLine || underline;
        //        textFormat.UnderlineType = rt.UnderLineType == ExcelUnderLineType.None ? underLineType : rt.UnderLineType;
        //        textFormat.SuperScript = rt.VerticalAlign == ExcelVerticalAlignmentFont.Superscript;
        //        textFormat.SubScript = rt.VerticalAlign == ExcelVerticalAlignmentFont.Subscript;
        //        textFormat.FontColor = rt.Color;
        //        textFormat.SubFamily = FontSubFamily.Regular;
        //        if (textFormat.Bold)
        //        {
        //            textFormat.SubFamily = FontSubFamily.Bold;
        //            if (textFormat.Italic)
        //            {
        //                textFormat.SubFamily = FontSubFamily.BoldItalic;
        //            }
        //        }
        //        else if (textFormat.Italic)
        //        {
        //            textFormat.SubFamily = FontSubFamily.Italic;
        //        }
        //        textFormats.Add(textFormat);
        //        dictionaries.AddFont(pageSettings, textFormat.FullFontName, textFormat.SubFamily, textFormat.Text);
        //    }
        //    return textFormats;
        //}

        public static PdfCellAlignmentData GetAlignmentData(PdfHeaderFooter headerFooter)
        {
            var contentAlignment = new PdfCellAlignmentData();
            switch (headerFooter.Alignment)
            {
                case HeaderFooterAlignment.Left:
                    contentAlignment.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    break;
                case HeaderFooterAlignment.Center:
                    contentAlignment.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    break;
                case HeaderFooterAlignment.Right:
                    contentAlignment.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                    break;
            }
            return contentAlignment;
        }

        public static PdfHeaderFooter GetTextFormats(PdfPageSettings pageSettings, PdfDictionaries dictionaries, ExcelWorksheet ws, ExcelHeaderFooterTextCollection textCollection, HeaderFooterType type, HeaderFooterAlignment alignment, HeaderFooterSection section)
        {
            if (textCollection == null || textCollection.Count <= 1) return null;
            var ns = ws.Workbook.Styles.GetNormalStyle();
            var textFragments = new List<TextFragment>();
            List<int> NumberOfPagesIndexes = new List<int>();
            List<int> PageNumberIndexes = new List<int>();
            for (int i = 1; i < textCollection.Count; i++)
            {
                var hf = textCollection[i];
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
                                      //(textFrag.RichTextOptions.Bold ? MeasurementFontStyles.Bold : 0) |
                                      //(textFrag.RichTextOptions.Italic ? MeasurementFontStyles.Italic : 0) |
                                      //(textFrag.RichTextOptions.UnderlineType != 12 ? MeasurementFontStyles.Underline : 0) |
                                      //(textFrag.RichTextOptions.StrikeType > 1 ? MeasurementFontStyles.Strikeout : 0);

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
                        NumberOfPagesIndexes.Add(i-1);
                        break;
                    case ExcelHeaderFooterFormattingCodes.PageNumber:
                        text += "000";
                        PageNumberIndexes.Add(i-1);
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




                //var text = textCollection[i];
                //var ns = ws.Workbook.Styles.GetNormalStyle();
                //PdfTextFormat textFormat = new PdfTextFormat();
                //textFormat.FontName = string.IsNullOrEmpty(text.FontName) ? ns.Style.Font.Name : text.FontName;
                //textFormat.FontSize = text.FontSize == null ? ns.Style.Font.Size : (double)text.FontSize;
                //textFormat.Bold = text.Bold;
                //textFormat.Italic = text.Italic;
                //textFormat.Strike = text.Striketrough;
                //textFormat.SubScript = text.SubScript;
                //textFormat.SuperScript = text.SuperScript;
                //textFormat.Underline = text.Underline;
                //textFormat.UnderlineType = textFormat.Underline ? ExcelUnderLineType.Single : ExcelUnderLineType.None;
                //if (text.DoubleUnderline) textFormat.Underline = true;
                //textFormat.UnderlineType = text.DoubleUnderline ? ExcelUnderLineType.Double : textFormat.UnderlineType;
                //textFormat.FontColor = text.Color;
                //textFormat.SubFamily = FontSubFamily.Regular;
                //if (textFormat.Bold)
                //{
                //    textFormat.SubFamily = FontSubFamily.Bold;
                //    if (textFormat.Italic)
                //    {
                //        textFormat.SubFamily = FontSubFamily.BoldItalic;
                //    }
                //}
                //else if (textFormat.Italic)
                //{
                //    textFormat.SubFamily = FontSubFamily.Italic;
                //}
                //switch (text.FormatCode)
                //{
                //    case ExcelHeaderFooterFormattingCodes.SheetName:
                //        textFormat.Text += ws.Name;
                //        break;
                //    case ExcelHeaderFooterFormattingCodes.CurrentDate:
                //        textFormat.Text += DateTime.Now.ToString($"yyyy-MM-dd");
                //        break;
                //    case ExcelHeaderFooterFormattingCodes.FileName:
                //        textFormat.Text += ws._package.File.Name;
                //        break;
                //    case ExcelHeaderFooterFormattingCodes.NumberOfPages:
                //    case ExcelHeaderFooterFormattingCodes.PageNumber:
                //        textFormat.Text += "000";
                //        containsPageNumber = true;
                //        break;
                //    case ExcelHeaderFooterFormattingCodes.CurrentTime:
                //        textFormat.Text += DateTime.Now.ToString("HH:mm");
                //        break;
                //    case ExcelHeaderFooterFormattingCodes.FilePath:
                //        textFormat.Text += ws._package.File.Directory.FullName + "\\";
                //        break;
                //    default:
                //        textFormat.Text += text.Text;
                //        break;
                //}
                //textFormats.Add(textFormat);
                //dictionaries.AddFont(pageSettings, textFormat.FullFontName, textFormat.SubFamily, textFormat.Text);
                //if (containsPageNumber) dictionaries.AddFont(pageSettings, textFormat.FullFontName, textFormat.SubFamily, "1234567890");
            }
            return new PdfHeaderFooter(textFragments, PageNumberIndexes, NumberOfPagesIndexes, type, alignment, section);
        }


        /// <summary>
        /// Check if we need to add additional columns to accomodate text that is not wrapped and overlaps other cell.s
        /// </summary>
        /// <param name="ws">The worksheet to check.</param>
        /// <returns>The number of columns to add.</returns>
        public static int AddColumnsForNonWrappedText(PdfPageSettings pageSettings, ExcelWorksheet ws, PdfWorksheet pdfSheet)
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
    }
}
