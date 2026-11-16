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
using EPPlus.Export.Pdf.DocumentObjects;
using EPPlus.Export.Pdf.Resources;
using EPPlus.Export.Pdf.Settings;
using EPPlus.Graphics.Units;
using OfficeOpenXml.CellPictures;
using OfficeOpenXml.ConditionalFormatting;
using OfficeOpenXml.Export.PdfExport.Data;
using OfficeOpenXml.RichData.Structures.Constants;
using OfficeOpenXml.Style.Table;
using OfficeOpenXml.Table;
using System;
using System.Collections.Generic;


namespace OfficeOpenXml.Export.PdfExport.TextMapping
{
    internal class PdfTextMap
    {
        readonly PdfPageSettings _pageSettings;
        readonly PdfDictionaries _dictionaries;
        readonly PdfWorksheet _pdfSheet;
        readonly ExcelWorksheet _worksheet;
        readonly double _zeroCharWidth;
        readonly CellPicturesManager _cellPictures;
        readonly int _addedColumns;
        readonly Dictionary<ExcelTable, ExcelTableNamedStyle> _tableStyleCache = new Dictionary<ExcelTable, ExcelTableNamedStyle>();
        readonly List<string> _checkedMergedCells = new List<string>();
        readonly PdfCellCollection _map;

        PdfRange _range;

        private PdfTextMap(PdfPageSettings pageSettings, PdfDictionaries dictionaries, PdfWorksheet pdfSheet, PdfRange range)
        {
            _pageSettings = pageSettings;
            _dictionaries = dictionaries;
            _pdfSheet = pdfSheet;
            _range = range;
            _worksheet = range.Range.Worksheet;
            _zeroCharWidth = pdfSheet.ZeroCharWidth = PdfWorksheet.GetThemeFont0Width(_worksheet, pageSettings.FontEngine);
            _cellPictures = new CellPicturesManager(_worksheet);
            _addedColumns = range.ExtendColumns ? AddColumnsForNonWrappedText(pageSettings, _worksheet, pdfSheet) : 0;
            _map = new PdfCellCollection(range.Range._fromRow, range.Range._toRow, range.Range._fromCol, range.Range._toCol + _addedColumns);
            _pdfSheet.ToRow = _pdfSheet.ToRow < range.Range._toRow ? range.Range._toRow : _pdfSheet.ToRow;
        }

        public static PdfCellCollection SetTextMap(PdfPageSettings pageSettings, PdfDictionaries dictionaries, PdfWorksheet pdfSheet, ref PdfRange pdfRange)
        {
            var builder = new PdfTextMap(pageSettings, dictionaries, pdfSheet, pdfRange);
            var map = builder.Build();
            pdfRange = builder._range;
            return map;
        }

        private PdfCellCollection Build()
        {
            for (int row = _range.Range._fromRow; row <= _range.Range._toRow; row++)
            {
                MapRow(row, firstColumnRun: row == _range.Range._fromRow);
            }
            _worksheet.ConditionalFormatting.ClearTempExportCacheForAllCFs();
            PdfBorderReconciler.ReconcileSharedBorders(_map);
            return _map;
        }

        private void MapRow(int row, bool firstColumnRun)
        {
            bool hiddenRow = ComputeRowHeight(row);
            for (int col = _range.Range._fromCol; col <= _range.Range._toCol + _addedColumns; col++)
            {
                MapCell(row, col, hiddenRow, firstColumnRun);
            }
        }

        private bool ComputeRowHeight(int row)
        {
            var hiddenRow = _worksheet.Row(row).Hidden;
            var r = (RowInternal)_worksheet.GetValueInner(row, 0);
            bool usesDefaultValue;
            double height;
            if (r == null || r.Height < 0)
            {
                usesDefaultValue = true;
                height = _worksheet.DefaultRowHeight;
            }
            else
            {
                usesDefaultValue = !r.CustomHeight;
                height = UnitConversion.ExcelRowHeightToPoints(r.Height);
            }
            _range.TotalHeight += hiddenRow ? 0d : height;
            _range.RowHeights.Add(new RowHeight { Height = hiddenRow ? 0d : height, UsesDefaultValue = usesDefaultValue });
            return hiddenRow;
        }

        private double ComputeColumnWidth(int col, bool firstColumnRun, out bool hiddenCol)
        {
            hiddenCol = _worksheet.Column(col).Hidden;
            var width = UnitConversion.ExcelColumnWidthToPoints(_worksheet.Column(col).Width, _zeroCharWidth);
            if (firstColumnRun)
            {
                _range.TotalWidth += hiddenCol ? 0d : width;
                _range.ColWidths.Add(hiddenCol ? 0d : width);
            }
            return width;
        }

        private void MapCell(int row, int col, bool hiddenRow, bool firstColumnRun)
        {
            var width = ComputeColumnWidth(col, firstColumnRun, out var hiddenCol);
            var tempMap = new PdfCell();
            tempMap.Hidden = hiddenRow || hiddenCol;
            tempMap.ColumnWidth = tempMap.Width = hiddenCol ? 0d : width;
            var cell = _worksheet.Cells[row, col];
            tempMap.Name = cell.Address;
            if (cell.Merge)
            {
                HandleMergedCell(cell, tempMap);
            }
            var cellStyle = new PdfCellStyle();
            PdfBorderStyleResolver.GetBorderStyles(cell, cellStyle, tempMap, _tableStyleCache);
            if (tempMap.Main == null)
            {
                BuildCellContent(cell, row, col, tempMap, cellStyle);
            }
            tempMap.CellStyle = cellStyle;
            _map[row, col] = tempMap;
            CollectComments(cell);
        }

        private void BuildCellContent(ExcelRange cell, int row, int col, PdfCell tempMap, PdfCellStyle cellStyle)
        {
            PdfFillStyleResolver.GetFillStyles(cell, cellStyle, _tableStyleCache);
            PdfFontStyleResolver.GetFontStyle(cell, cellStyle, _tableStyleCache);
            tempMap.ContentAligmnet = PdfCellAlignmentResolver.GetContentAlignment(cell);
            bool hideText = ApplyConditionalFormatVisuals(cell, tempMap, cellStyle);
            var cellPicture = _cellPictures.GetCellPicture(row, col) ?? _cellPictures.GetCellPicture(row, col, StructureTypes.WebImage);
            if (cellPicture != null)
            {
                LoadCellPicture(cellPicture, tempMap);
            }

            if (!string.IsNullOrEmpty(cell.Text) && !hideText)
            {
                tempMap.Text = cell.Text;
                tempMap.TextFragments = PdfTextFragmentBuilder.GetTextFragments(_pageSettings, _dictionaries, cell, cellStyle);
            }

            ResolveCellHyperlink(cell, tempMap);
        }

        private bool ApplyConditionalFormatVisuals(ExcelRange cell, PdfCell tempMap, PdfCellStyle cellStyle)
        {
            var dataBarRule = PdfDataBarResolver.GetDataBarRule(cell);
            PdfDataBarResolver.GetDataBarStyle(cell, cellStyle, dataBarRule);
            var iconRule = PdfDataBarResolver.GetIconSetRule(cell);
            if (iconRule != null)
            {
                tempMap.ContentAligmnet.IconSize = cell.Style.Font.Size * 1.2d;
                try
                {
                    var icon = iconRule.GetCustomIcon(new ExcelAddress(cell.Address));
                    if (icon != null && icon != eExcelconditionalFormattingCustomIcon.NoIcon)
                    {
                        cellStyle.Icon = icon;
                    }
                }
                catch (Exception)
                {
                    cellStyle.Icon = null;
                    tempMap.ContentAligmnet.IconSize = 0d;
                }
            }
            return (dataBarRule != null && !dataBarRule.ShowValue) || (iconRule != null && !iconRule.ShowValue);
        }

        private void ResolveCellHyperlink(ExcelRange cell, PdfCell tempMap)
        {
            var hl = cell.Hyperlink;
            string screenTip = (hl as ExcelHyperLink)?.ToolTip;
            if (hl is ExcelHyperLink eh && !string.IsNullOrEmpty(eh.ReferenceAddress))
            {
                if (PdfHyperlinkResolver.TryResolveInternalLink(_worksheet, eh.ReferenceAddress, out var ts, out var tr, out var tc))
                {
                    tempMap.HyperlinkTargetSheet = ts;
                    tempMap.HyperlinkTargetRow = tr;
                    tempMap.HyperlinkTargetCol = tc;
                    tempMap.HyperlinkToolTip = !string.IsNullOrEmpty(screenTip) ? screenTip : eh.ReferenceAddress;
                }
            }
            else if (hl != null)
            {
                tempMap.HyperlinkUri = PdfHyperlinkResolver.NormalizeExternalTarget(hl.OriginalString);
                if (tempMap.HyperlinkUri != null)
                {
                    tempMap.HyperlinkToolTip = !string.IsNullOrEmpty(screenTip) ? screenTip : hl.OriginalString;
                }
            }
            else if (PdfHyperlinkResolver.TryGetFormulaHyperlink(cell, out var fUri, out var fInternal))
            {
                if (fUri != null)
                {
                    tempMap.HyperlinkUri = fUri;
                    tempMap.HyperlinkToolTip = fUri;
                }
                else if (fInternal != null && PdfHyperlinkResolver.TryResolveInternalLink(_worksheet, fInternal, out var ts, out var tr, out var tc))
                {
                    tempMap.HyperlinkTargetSheet = ts;
                    tempMap.HyperlinkTargetRow = tr;
                    tempMap.HyperlinkTargetCol = tc;
                    tempMap.HyperlinkToolTip = fInternal;
                }
            }
        }

        private void CollectComments(ExcelRange cell)
        {
            if (_pageSettings.CommentsAndNotes == CommentsAndNotes.None) return;
            if (cell.Comment != null && cell.ThreadedComment == null)
            {
                _pdfSheet.CommentsAndNotesCollections.Add(cell.Address, new PdfCommentsAndNotes(cell.Comment));
            }
            if (cell.ThreadedComment != null)
            {
                _pdfSheet.CommentsAndNotesCollections.Add(cell.Address, new PdfCommentsAndNotes(cell.ThreadedComment));
                PdfCommentsAndNotes.HasThreadedComment = true;
            }
        }

        private void HandleMergedCell(ExcelRange cell, PdfCell tempMap)
        {
            var worksheet = cell.Worksheet;
            string mergeAddress = worksheet.MergedCells[cell.Start.Row, cell.Start.Column];
            ExcelAddressBase address = new ExcelAddressBase(mergeAddress);
            if (!_checkedMergedCells.Contains(mergeAddress))
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
                    totalWidth += UnitConversion.ExcelColumnWidthToPoints(worksheet.Column(l).Width, _zeroCharWidth);
                }
                _checkedMergedCells.Add(mergeAddress);
                tempMap.Width = totalWidth;
                tempMap.Height = totalHeight;
                tempMap.Main = null;
            }
            else
            {
                tempMap.Main = _map[address._fromRow, address._fromCol];
                if (tempMap.Main == null)
                {
                    var main = worksheet.Cells[address._fromRow, address._fromCol];
                    PdfCell mainCell = new PdfCell();
                    var cellStyle = new PdfCellStyle();
                    PdfBorderStyleResolver.GetBorderStyles(main, cellStyle, mainCell, _tableStyleCache);
                    PdfFillStyleResolver.GetFillStyles(main, cellStyle, _tableStyleCache);
                    PdfFontStyleResolver.GetFontStyle(main, cellStyle, _tableStyleCache);
                    mainCell.ContentAligmnet = PdfCellAlignmentResolver.GetContentAlignment(main);
                    if (!string.IsNullOrEmpty(main.Text))
                    {
                        mainCell.Text = main.Text;
                        mainCell.TextFragments = PdfTextFragmentBuilder.GetTextFragments(_pageSettings, _dictionaries, main, cellStyle);
                    }
                    mainCell.CellStyle = cellStyle;
                    tempMap.Main = mainCell;
                }
            }
            tempMap.MergedAddress = address;
            tempMap.Name = tempMap.Name + " ; " + address.ToString();
            tempMap.Merged = true;
        }

        private static void LoadCellPicture(ExcelCellPicture picture, PdfCell target)
        {
            try
            {
                var bytes = picture.GetImageBytes();
                if (bytes == null || !PdfImageXObject.CanEmbed(bytes)) return;
                var bounds = picture.GetImage().Bounds;
                if (bounds.Width <= 0 || bounds.Height <= 0) return;
                target.CellPictureBytes = bytes;
                target.CellPicturePixelWidth = bounds.Width;
                target.CellPicturePixelHeight = bounds.Height;
            }
            catch
            {
                // Corrupt or unreadable image — leave the cell blank rather than throwing.
            }
        }

        /// <summary>
        /// Check if we need to add additional columns to accomodate text that is not wrapped and overlaps other cell.s
        /// </summary>
        /// <param name="ws">The worksheet to check.</param>
        /// <returns>The number of columns to add.</returns>
        internal static int AddColumnsForNonWrappedText(PdfPageSettings pageSettings, ExcelWorksheet ws, PdfWorksheet pdfSheet)
        {
            if (ws.Dimension == null) return 0; //If dimension is null, there are no cells to check, so return 0.
            int columnsToAdd = 0;
            var catalog = new EPPlusToPdfWriter();
            var lastColumn = ws.Dimension.End.Column;
            ExcelRangeBase lastColumnRange = ws.Cells[1, lastColumn, ws.Dimension.End.Row, lastColumn];
            var cc = catalog.GetCellCollectionFromRange(pageSettings, lastColumnRange);
            double textLength = 0;
            for (int i = cc.FromRow; i <= cc.ToRow; i++)
            {
                var lastColCell = cc[i, cc.FromColumn];
                if (lastColCell == null)
                {
                    continue;
                }
                if (lastColCell.ContentAligmnet != null && lastColCell.ContentAligmnet.WrapText)
                {
                    continue;
                }
                if (lastColCell.Merged)
                {
                    continue;
                }
                if (lastColCell.TotalTextLength > textLength)
                {
                    textLength = lastColCell.TotalTextLength;
                }
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