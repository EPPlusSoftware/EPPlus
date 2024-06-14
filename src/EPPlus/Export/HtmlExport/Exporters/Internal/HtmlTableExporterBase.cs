/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  6/4/2022         EPPlus Software AB           ExcelTable Html Export
 *************************************************************************************************/
using OfficeOpenXml.Export.HtmlExport.Exporters.Internal;
using OfficeOpenXml.Table;
using OfficeOpenXml.Export.HtmlExport.Settings;
using OfficeOpenXml.Utils;
using System.Collections.Generic;
using System.Linq;
using OfficeOpenXml.Export.HtmlExport.HtmlCollections;

namespace OfficeOpenXml.Export.HtmlExport.Exporters
{
    internal abstract class HtmlTableExporterBase : HtmlExporterBaseInternal
    {
        internal HtmlTableExporterBase
            (HtmlTableExportSettings settings, ExcelTable table, ExcelRangeBase range) : base(settings, range)
        {
            Require.Argument(table).IsNotNull("table");
            _table = table;
            _tableExportSettings = settings;

            LoadRangeImages(new List<ExcelRangeBase>() { table.Range });
        }

        protected readonly ExcelTable _table;
        protected readonly HtmlTableExportSettings _tableExportSettings;


        protected override void AddTableData(ExcelTable table, HTMLElement th, int col)
        {
            if (table != null)
            {
                if (Settings.Accessibility.TableSettings.AddAccessibilityAttributes && !string.IsNullOrEmpty(Settings.Accessibility.TableSettings.TableHeaderCellRole))
                {
                    th.AddAttribute("role", Settings.Accessibility.TableSettings.TableHeaderCellRole);
                    if (!table.ShowFirstColumn && !table.ShowLastColumn)
                    {
                        th.AddAttribute("scope", "col");
                    }
                    if (table.SortState != null && !table.SortState.ColumnSort && table.SortState.SortConditions.Any())
                    {
                        var firstCondition = table.SortState.SortConditions.First();
                        if (firstCondition != null && !string.IsNullOrEmpty(firstCondition.Ref))
                        {
                            var addr = new ExcelAddress(firstCondition.Ref);
                            var sortedCol = addr._fromCol;
                            if (col == sortedCol)
                            {
                                th.AddAttribute("aria-sort", firstCondition.Descending ? "descending" : "ascending");
                            }
                        }
                    }
                }
            }
        }

        private void LoadVisibleColumns()
        {
            _columns = new List<int>();
            var r = _table.Range;
            for (int col = r._fromCol; col <= r._toCol; col++)
            {
                var c = _table.WorkSheet.GetColumn(col);
                if (c == null || c.Hidden == false && c.Width > 0)
                {
                    _columns.Add(col);
                }
            }
        }

        protected HTMLElement GenerateHtml()
        {
            GetDataTypes(_table.Address, _table);

            var htmlTable = new HTMLElement(HtmlElements.Table);

            HtmlExportTableUtil.AddClassesAttributes(htmlTable, _table, _tableExportSettings);
            AddTableAccessibilityAttributes(Settings.Accessibility, htmlTable);

            LoadVisibleColumns();
            if (Settings.SetColumnWidth || Settings.HorizontalAlignmentWhenGeneral == eHtmlGeneralAlignmentHandling.ColumnDataType)
            {
                SetColumnGroup(htmlTable, _table.Range, Settings, false);
            }

            if (_table.ShowHeader)
            {
                AddHeaderRow(htmlTable);
            }

            AddTableRows(htmlTable);

            if (_table.ShowTotal)
            {
                AddTotalRow(htmlTable);
            }

            return htmlTable;
        }

        private void AddTableRows(HTMLElement htmlTable)
        {
            var row = _table.ShowHeader ? _table.Address._fromRow + 1 : _table.Address._fromRow;
            var endRow = _table.ShowTotal ? _table.Address._toRow - 1 : _table.Address._toRow;

            var body = GetTableBody(_table.Range, row, endRow);
            htmlTable.AddChildElement(body);
        }

        private void AddHeaderRow(HTMLElement table)
        {
            table.AddChildElement(GetThead(_table.Range));
        }

        private void AddTotalRow(HTMLElement table)
        {
            // table header row
            var tFoot = new HTMLElement(HtmlElements.TFoot);

            var rowIndex = _table.Address._toRow;
            if (Settings.Accessibility.TableSettings.AddAccessibilityAttributes && !string.IsNullOrEmpty(Settings.Accessibility.TableSettings.TfootRole))
            {
                tFoot.AddAttribute("role", Settings.Accessibility.TableSettings.TfootRole);
            }

            var row = new HTMLElement(HtmlElements.TableRow);

            if (Settings.Accessibility.TableSettings.AddAccessibilityAttributes)
            {
                row.AddAttribute("role", "row");
                row.AddAttribute("scope", "row");
            }
            if (Settings.SetRowHeight) AddRowHeightStyle(row, _table.Range, rowIndex, Settings.StyleClassPrefix, false);

            var address = _table.Address;
            HtmlImage image = null;
            foreach (var col in _columns)
            {
                var tblData = new HTMLElement(HtmlElements.TableData);

                var cell = _table.WorkSheet.Cells[rowIndex, col];
                if (Settings.Accessibility.TableSettings.AddAccessibilityAttributes)
                {
                    tblData.AddAttribute("role", "cell");
                }
                GetClassData(tblData, true, image, cell, Settings, _exporterContext, out HTMLElement contentElement);

                AddImage(contentElement, Settings, image, cell.Value);

                contentElement.Content = GetCellText(cell, Settings);

                row.AddChildElement(tblData);
            }
            tFoot.AddChildElement(row);
            table.AddChildElement(tFoot);
        }
    }
}
