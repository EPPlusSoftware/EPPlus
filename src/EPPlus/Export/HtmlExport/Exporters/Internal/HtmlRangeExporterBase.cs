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
using OfficeOpenXml.Core;
using OfficeOpenXml.Drawing.Interfaces;
using OfficeOpenXml.Export.HtmlExport.HtmlCollections;
using OfficeOpenXml.Export.HtmlExport.Parsers;
using OfficeOpenXml.Table;
using System;
using System.Collections.Generic;
using System.Runtime;
using System.Linq;
using System.Xml.Linq;
using OfficeOpenXml.Core.CellStore;

namespace OfficeOpenXml.Export.HtmlExport.Exporters.Internal
{
    internal abstract class HtmlRangeExporterBase : HtmlExporterBaseInternal
    {
        internal HtmlRangeExporterBase(HtmlRangeExportSettings settings, ExcelRangeBase range) : base(settings, range)
        {
            _settings = settings;
        }

        internal HtmlRangeExporterBase(HtmlRangeExportSettings settings, EPPlusReadOnlyList<ExcelRangeBase> ranges) : base(settings, ranges)
        {
            _settings = settings;
        }

        protected readonly HtmlRangeExportSettings _settings;

        protected HTMLElement GenerateHTML(int rangeIndex, ExcelHtmlOverrideExportSettings overrideSettings = null)
        {
            ValidateRangeIndex(rangeIndex);
            _mergedCells.Clear();
            var range = _ranges[rangeIndex];
            GetDataTypes(range, _settings);

            ExcelTable table = null;
            if (Settings.TableStyle != eHtmlRangeTableInclude.Exclude)
            {
                table = range.GetTable();
            }

            var tableId = GetTableId(rangeIndex, overrideSettings);
            var additionalClassNames = GetAdditionalClassNames(overrideSettings);
            var accessibilitySettings = GetAccessibilitySettings(overrideSettings);
            var headerRows = overrideSettings != null ? overrideSettings.HeaderRows : _settings.HeaderRows;
            var headers = overrideSettings != null ? overrideSettings.Headers : _settings.Headers;

            var htmlTable = new HTMLElement(HtmlElements.Table);

            AddClassesAttributes(htmlTable, table, tableId, additionalClassNames);
            AddTableAccessibilityAttributes(accessibilitySettings, htmlTable);

            LoadVisibleColumns(range);
            if (Settings.SetColumnWidth || Settings.HorizontalAlignmentWhenGeneral == eHtmlGeneralAlignmentHandling.ColumnDataType)
            {
                SetColumnGroup(htmlTable, range, Settings, IsMultiSheet);
            }

            if (_settings.HeaderRows > 0 || _settings.Headers.Count > 0)
            {
                AddHeaderRow(range, htmlTable, table, headers);
            }
            // table rows
            AddTableRows(htmlTable, range);

            return htmlTable;
        }

        private void AddTableRows(HTMLElement htmlTable, ExcelRangeBase range)
        {
            var row = range._fromRow + _settings.HeaderRows;

            var body = GetTableBody(range, row, range._toRow);
            htmlTable.AddChildElement(body);
        }

        private void AddHeaderRow(ExcelRangeBase range, HTMLElement element, ExcelTable table, List<string> headers)
        {
            if (table != null && table.ShowHeader == false) return;

            var thead = GetThead(range, headers);

            element.AddChildElement(thead);
        }

        protected override int GetHeaderRows(ExcelTable table)
        {
            int headerRows;

            if (table == null)
            {
                headerRows = _settings.HeaderRows == 0 ? 1 : _settings.HeaderRows;
            }
            else
            {
                headerRows = table.ShowHeader ? 1 : 0;
            }

            return headerRows;
        }
    }
}
