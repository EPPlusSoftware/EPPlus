using OfficeOpenXml.Core;
using OfficeOpenXml.Drawing.Interfaces;
using OfficeOpenXml.Export.HtmlExport.Accessibility;
using OfficeOpenXml.Table;
using OfficeOpenXml.Utils;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Export.HtmlExport.Exporters
{
    internal abstract class HtmlExporterBase : AbstractExporter
    {
        public HtmlExporterBase(HtmlExportSettings settings, ExcelRangeBase range)
        {
            Settings = settings;
            Require.Argument(range).IsNotNull("range");
            _ranges = new EPPlusReadOnlyList<ExcelRangeBase>();

            if (range.Addresses == null)
            {
                AddRange(range);
            }
            else
            {
                foreach (var address in range.Addresses)
                {
                    AddRange(range.Worksheet.Cells[address.Address]);
                }
            }

            LoadRangeImages(_ranges._list);
        }

        public HtmlExporterBase(HtmlExportSettings settings, ExcelRangeBase[] ranges)
        {
            Settings = settings;
            Require.Argument(ranges).IsNotNull("ranges");
            _ranges = new EPPlusReadOnlyList<ExcelRangeBase>();

            foreach (var range in ranges)
            {
                AddRange(range);
            }

            LoadRangeImages(_ranges._list);
        }

        public HtmlExporterBase(ExcelRangeBase[] ranges)
        {
            Require.Argument(ranges).IsNotNull("ranges");
            _ranges = new EPPlusReadOnlyList<ExcelRangeBase>();

            foreach (var range in ranges)
            {
                AddRange(range);
            }

            LoadRangeImages(_ranges._list);
        }

        protected List<int> _columns = new List<int>();
        protected readonly Dictionary<string, int> _styleCache = new Dictionary<string, int>();
        protected HtmlExportSettings Settings;
        protected readonly List<ExcelAddressBase> _mergedCells = new List<ExcelAddressBase>();

        protected void LoadVisibleColumns(ExcelRangeBase range)
        {
            var ws = range.Worksheet;
            _columns = new List<int>();
            for (int col = range._fromCol; col <= range._toCol; col++)
            {
                var c = ws.GetColumn(col);
                if (c == null || (c.Hidden == false && c.Width > 0))
                {
                    _columns.Add(col);
                }
            }
        }

        protected EPPlusReadOnlyList<ExcelRangeBase> _ranges = new EPPlusReadOnlyList<ExcelRangeBase>();

        private void AddRange(ExcelRangeBase range)
        {
            if (range.IsFullColumn && range.IsFullRow)
            {
                _ranges.Add(new ExcelRangeBase(range.Worksheet, range.Worksheet.Dimension.Address));
            }
            else
            {
                _ranges.Add(range);
            }
        }

        protected void ValidateRangeIndex(int rangeIndex)
        {
            if (rangeIndex < 0 || rangeIndex >= _ranges.Count)
            {
                throw new ArgumentOutOfRangeException(nameof(rangeIndex));
            }
        }

        protected string GetWorksheetClassName(string styleClassPrefix, string name, ExcelWorksheet ws, bool addWorksheetName)
        {
            if (addWorksheetName)
            {
                return styleClassPrefix + name + "-" + GetClassName(ws.Name, $"Sheet{ws.PositionId}");
            }
            else
            {
                return styleClassPrefix + name;
            }
        }

        internal bool HandleHiddenRow(EpplusHtmlWriter writer, ExcelWorksheet ws, HtmlExportSettings Settings, ref int row)
        {
            if (Settings.HiddenRows != eHiddenState.Include)
            {
                var r = ws.Row(row);
                if (r.Hidden || r.Height == 0)
                {
                    if (Settings.HiddenRows == eHiddenState.IncludeButHide)
                    {
                        writer.AddAttribute("class", $"{Settings.StyleClassPrefix}hidden");
                    }
                    else
                    {
                        row++;
                        return true;
                    }
                }
            }

            return false;
        }

        internal void AddRowHeightStyle(EpplusHtmlWriter writer, ExcelRangeBase range, int row, string styleClassPrefix, bool isMultiSheet)
        {
            var r = range.Worksheet._values.GetValue(row, 0);
            if (r._value is RowInternal rowInternal)
            {
                if (rowInternal.Height != -1 && rowInternal.Height != range.Worksheet.DefaultRowHeight)
                {
                    writer.AddAttribute("style", $"height:{rowInternal.Height}pt");
                    return;
                }
            }

            var clsName = GetWorksheetClassName(styleClassPrefix, "drh", range.Worksheet, isMultiSheet);
            writer.AddAttribute("class", clsName); //Default row height
        }

        protected string GetPictureName(HtmlImage p)
        {
            var hash = ((IPictureContainer)p.Picture).ImageHash;
            var fi = new FileInfo(p.Picture.Part.Uri.OriginalString);
            var name = fi.Name.Substring(0, fi.Name.Length - fi.Extension.Length);

            return GetClassName(name, hash);
        }

        protected bool InMergeCellSpan(int row, int col)
        {
            for (int i = 0; i < _mergedCells.Count; i++)
            {
                var adr = _mergedCells[i];
                if (adr._toRow < row || (adr._toRow == row && adr._toCol < col))
                {
                    _mergedCells.RemoveAt(i);
                    i--;
                }
                else
                {
                    if (row >= adr._fromRow && row <= adr._toRow &&
                       col >= adr._fromCol && col <= adr._toCol)
                    {
                        return true;
                    }
                }
            }
            return false;
        }

        protected void SetColRowSpan(ExcelRangeBase range, EpplusHtmlWriter writer, ExcelRange cell)
        {
            if (cell.Merge)
            {
                var address = cell.Worksheet.MergedCells[cell._fromRow, cell._fromCol];
                if (address != null)
                {
                    var ma = new ExcelAddressBase(address);
                    bool added = false;
                    //ColSpan
                    if (ma._fromCol == cell._fromCol || range._fromCol == cell._fromCol)
                    {
                        var maxCol = Math.Min(ma._toCol, range._toCol);
                        var colSpan = maxCol - ma._fromCol + 1;
                        if (colSpan > 1)
                        {
                            writer.AddAttribute("colspan", colSpan.ToString(CultureInfo.InvariantCulture));
                        }
                        _mergedCells.Add(ma);
                        added = true;
                    }
                    //RowSpan
                    if (ma._fromRow == cell._fromRow || range._fromRow == cell._fromRow)
                    {
                        var maxRow = Math.Min(ma._toRow, range._toRow);
                        var rowSpan = maxRow - ma._fromRow + 1;
                        if (rowSpan > 1)
                        {
                            writer.AddAttribute("rowspan", rowSpan.ToString(CultureInfo.InvariantCulture));
                        }
                        if (added == false) _mergedCells.Add(ma);
                    }
                }
            }
        }

        protected void GetDataTypes(ExcelRangeBase range, HtmlRangeExportSettings settings)
        {
            if (range._fromRow + settings.HeaderRows > ExcelPackage.MaxRows)
            {
                throw new InvalidOperationException("Range From Row + Header rows is out of bounds");
            }

            _dataTypes = new List<string>();
            for (int col = range._fromCol; col <= range._toCol; col++)
            {
                _dataTypes.Add(
                    ColumnDataTypeManager.GetColumnDataType(range.Worksheet, range, range._fromRow + settings.HeaderRows, col));
            }
        }
        bool? _isMultiSheet = null;
        protected bool IsMultiSheet
        {
            get
            {
                if (_isMultiSheet.HasValue == false)
                {
                    _isMultiSheet = _ranges.Select(x => x.Worksheet).Distinct().Count() > 1;
                }
                return _isMultiSheet.Value;
            }
        }

        protected void AddTableAccessibilityAttributes(AccessibilitySettings settings, EpplusHtmlWriter writer)
        {
            if (!settings.TableSettings.AddAccessibilityAttributes) return;
            if (!string.IsNullOrEmpty(settings.TableSettings.TableRole))
            {
                writer.AddAttribute("role", settings.TableSettings.TableRole);
            }
            if (!string.IsNullOrEmpty(settings.TableSettings.AriaLabel))
            {
                writer.AddAttribute(AriaAttributes.AriaLabel.AttributeName, settings.TableSettings.AriaLabel);
            }
            if (!string.IsNullOrEmpty(settings.TableSettings.AriaLabelledBy))
            {
                writer.AddAttribute(AriaAttributes.AriaLabelledBy.AttributeName, settings.TableSettings.AriaLabelledBy);
            }
            if (!string.IsNullOrEmpty(settings.TableSettings.AriaDescribedBy))
            {
                writer.AddAttribute(AriaAttributes.AriaDescribedBy.AttributeName, settings.TableSettings.AriaDescribedBy);
            }
        }

        protected string GetTableId(int index, ExcelHtmlOverrideExportSettings overrideSettings)
        {
            if (overrideSettings == null || string.IsNullOrEmpty(overrideSettings.TableId))
            {
                if (_ranges.Count > 1 && !string.IsNullOrEmpty(Settings.TableId))
                {
                    return Settings.TableId + index.ToString(CultureInfo.InvariantCulture);
                }
                return Settings.TableId;
            }
            return overrideSettings.TableId;
        }

        protected List<string> GetAdditionalClassNames(ExcelHtmlOverrideExportSettings overrideSettings)
        {
            if (overrideSettings == null || overrideSettings.AdditionalTableClassNames == null) return Settings.AdditionalTableClassNames;
            return overrideSettings.AdditionalTableClassNames;
        }

        protected AccessibilitySettings GetAccessibilitySettings(ExcelHtmlOverrideExportSettings overrideSettings)
        {
            if (overrideSettings == null || overrideSettings.Accessibility == null) return Settings.Accessibility;
            return overrideSettings.Accessibility;
        }

        protected void AddClassesAttributes(EpplusHtmlWriter writer, ExcelTable table, string tableId, List<string> additionalTableClassNames)
        {
            var tableClasses = TableClass;
            if (table != null)
            {
                tableClasses += " " + ExcelHtmlTableExporter.GetTableClasses(table); //Add classes for the table styles if the range corresponds to a table.
            }
            if (additionalTableClassNames != null && additionalTableClassNames.Count > 0)
            {
                foreach (var cls in additionalTableClassNames)
                {
                    tableClasses += $" {cls}";
                }
            }
            writer.AddAttribute(HtmlAttributes.Class, $"{tableClasses}");

            if (!string.IsNullOrEmpty(tableId))
            {
                writer.AddAttribute(HtmlAttributes.Id, tableId);
            }
        }
    }
}
