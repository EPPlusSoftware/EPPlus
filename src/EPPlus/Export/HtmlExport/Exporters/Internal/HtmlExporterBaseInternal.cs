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
using OfficeOpenXml.Export.HtmlExport.Accessibility;
using OfficeOpenXml.Export.HtmlExport.HtmlCollections;
using OfficeOpenXml.Export.HtmlExport.Parsers;
using OfficeOpenXml.Table;
using OfficeOpenXml.Utils;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;

namespace OfficeOpenXml.Export.HtmlExport.Exporters.Internal
{
    internal abstract class HtmlExporterBaseInternal : AbstractHtmlExporter
    {
        public HtmlExporterBaseInternal(HtmlExportSettings settings, ExcelRangeBase range)
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

        public HtmlExporterBaseInternal(HtmlExportSettings settings, EPPlusReadOnlyList<ExcelRangeBase> ranges)
        {
            Settings = settings;
            Require.Argument(ranges).IsNotNull("ranges");
            _ranges = ranges;
            //TODO: Fix support for all ranges
            LoadRangeImages(_ranges._list);
        }

        protected void SetColumnGroup(HTMLElement element, ExcelRangeBase _range, HtmlExportSettings settings, bool isMultiSheet)
        {
            var group = GetColumnGroup(_range, settings, isMultiSheet);
            element.AddChildElement(group);
        }

        HTMLElement GetColumnGroup(ExcelRangeBase _range, HtmlExportSettings settings, bool isMultiSheet)
        {
            var group = new HTMLElement("colgroup");

            var ws = _range.Worksheet;
            var mdw = _range.Worksheet.Workbook.MaxFontWidth;
            var defColWidth = ExcelColumn.ColumnWidthToPixels(Convert.ToDecimal(ws.DefaultColWidth), mdw);

            foreach (var c in _columns)
            {
                var element = new HTMLElement("col");
                if (settings.SetColumnWidth)
                {
                    double width = ws.GetColumnWidthPixels(c - 1, mdw);
                    if (width == defColWidth)
                    {
                        var clsName = HtmlExportTableUtil.GetWorksheetClassName(settings.StyleClassPrefix, "dcw", ws, isMultiSheet);
                        element.AddAttribute("class", clsName);
                    }
                    else
                    {
                        element.AddAttribute("style", $"width:{width}px");
                    }
                }
                if (settings.HorizontalAlignmentWhenGeneral == eHtmlGeneralAlignmentHandling.ColumnDataType)
                {
                    element.AddAttribute("class", $"{TableClass}-ar");
                }
                element.AddAttribute("span", "1");

                group.AddChildElement(element);
            }
            return group;
        }

        protected HTMLElement GetThead(ExcelRangeBase range, List<string> headers = null)
        {
            var thead = CreateThead();
            ExcelTable table = null;
            if (Settings.TableStyle != eHtmlRangeTableInclude.Exclude)
            {
                table = range.GetTable();
            }

            int headerRows = GetHeaderRows(table);

            for (int i = 0; i < headerRows; i++)
            {
                var row = range._fromRow + i;
                var rowElement = CreateRow(range, row);

                ExcelWorksheet worksheet = range.Worksheet;
                HtmlImage image = null;
                foreach (var col in _columns)
                {
                    if (InMergeCellSpan(row, col)) continue;

                    var th = new HTMLElement(HtmlElements.TableHeader);
                    var cell = worksheet.Cells[row, col];
                    if (Settings.RenderDataTypes)
                    {
                        th.AddAttribute("data-datatype", _dataTypes[col - range._fromCol]);
                    }

                    SetColRowSpan(range, th, cell);

                    HTMLElement contentElement;

                    if (Settings.IncludeCssClassNames)
                    {
                        GetClassData(th, true, image, cell, Settings, _exporterContext, out contentElement, true);
                    }
                    else
                    {
                        contentElement = th;
                    }


                    AddTableData(table, contentElement, col);

                    if (Settings.Pictures.Include == ePictureInclude.Include)
                    {
                        image = GetImage(cell.Worksheet.PositionId, cell._fromRow, cell._fromCol);
                    }

                    AddImage(contentElement, Settings, image, cell.Value);

                    if (headerRows > 0 || table != null)
                    {
                        if (cell.Hyperlink == null)
                        {
                            contentElement.Content = GetCellText(cell, Settings);
                        }
                        else
                        {
                            AddHyperlink(contentElement, cell, Settings);
                        }
                    }
                    else if (headers.Count < col)
                    {
                        contentElement.Content = headers[col];
                    }

                    rowElement.AddChildElement(th);
                }
                thead.AddChildElement(rowElement);
            }
            return thead;
        }

        HTMLElement CreateThead()
        {
            var thead = new HTMLElement(HtmlElements.Thead);

            if (Settings.Accessibility.TableSettings.AddAccessibilityAttributes && !string.IsNullOrEmpty(Settings.Accessibility.TableSettings.TheadRole))
            {
                thead.AddAttribute("role", Settings.Accessibility.TableSettings.TheadRole);
            }

            return thead;
        }

        HTMLElement CreateRow(ExcelRangeBase range, int row = 0)
        {
            var tr = new HTMLElement(HtmlElements.TableRow);
            if (Settings.Accessibility.TableSettings.AddAccessibilityAttributes)
            {
                tr.AddAttribute("role", "row");
            }

            if (Settings.SetRowHeight) AddRowHeightStyle(tr, range, row, Settings.StyleClassPrefix, IsMultiSheet);

            return tr;
        }

        protected virtual int GetHeaderRows(ExcelTable table)
        {
            return 1;
        }

        protected virtual void AddTableData(ExcelTable table, HTMLElement th, int col)
        {
        }

        protected HTMLElement GetTableBody(ExcelRangeBase range, int row, int endRow)
        {
            var tBody = new HTMLElement(HtmlElements.Tbody);
            if (Settings.Accessibility.TableSettings.AddAccessibilityAttributes && !string.IsNullOrEmpty(Settings.Accessibility.TableSettings.TbodyRole))
            {
                tBody.AddAttribute("role", Settings.Accessibility.TableSettings.TbodyRole);
            }

            var table = range.GetTable();

            var ws = range.Worksheet;
            HtmlImage image = null;
            bool hasFooter = table != null && table.ShowTotal;
            while (row <= endRow)
            {
                EpplusHtmlAttribute attribute = null;
                if (HandleHiddenRow(attribute, range.Worksheet, Settings, ref row))
                {
                    continue; //The row is hidden and should not be included.
                }

                HTMLElement tFoot = null;
                if (hasFooter && row == endRow)
                {
                    tFoot = new HTMLElement(HtmlElements.TFoot);
                    if (attribute != null) { tFoot.AddAttribute(attribute.AttributeName, attribute.Value); }
                    attribute = null;
                }

                var tr = new HTMLElement(HtmlElements.TableRow);

                if (attribute != null) { tr.AddAttribute(attribute.AttributeName, attribute.Value); }

                if (Settings.Accessibility.TableSettings.AddAccessibilityAttributes)
                {
                    tr.AddAttribute("role", "row");

                    if (table == null || !table.ShowFirstColumn && !table.ShowLastColumn)
                    {
                        tr.AddAttribute("scope", "row");
                    }
                }

                if (Settings.SetRowHeight) AddRowHeightStyle(tr, range, row, Settings.StyleClassPrefix, IsMultiSheet);

                foreach (var col in _columns)
                {
                    if (InMergeCellSpan(row, col)) continue;
                    var colIx = col - range._fromCol;
                    var cell = ws.Cells[row, col];

                    var dataType = HtmlRawDataProvider.GetHtmlDataTypeFromValue(cell.Value);

                    var tblData = new HTMLElement(HtmlElements.TableData);

                    SetColRowSpan(range, tblData, cell);

                    if (Settings.Pictures.Include == ePictureInclude.Include)
                    { 
                        image = GetImage(cell.Worksheet.PositionId, cell._fromRow, cell._fromCol);
                    }

                    if (cell.Hyperlink == null)
                    {
                        var addRowScope = table == null ? false : table.ShowFirstColumn && col == table.Address._fromCol || table.ShowLastColumn && col == table.Address._toCol;
                        AddTableDataFromCell(cell, dataType, tblData, Settings, addRowScope, image, _exporterContext);
                    }
                    else
                    {
                        GetClassData(tblData, table != null, image, cell, Settings, _exporterContext, out HTMLElement contentElement);

                        AddImage(contentElement, Settings, image, cell.Value);
                        AddHyperlink(contentElement, cell, Settings);
                    }
                    tr.AddChildElement(tblData);
                }

                tBody.AddChildElement(tr);

                if (tFoot != null)
                {
                    tBody.AddChildElement(tFoot);
                }
                row++;
            }

            range.Worksheet.ConditionalFormatting.ClearTempExportCacheForAllCFs();
            return tBody;
        }

        /// <summary>
        /// Renders a hyperlink
        /// </summary>
        /// <param name="element"></param>
        /// <param name="cell"></param>
        /// <param name="settings"></param>
        protected void AddHyperlink(HTMLElement element, ExcelRangeBase cell, HtmlExportSettings settings)
        {
            if (cell.Hyperlink is ExcelHyperLink eurl)
            {
                if (string.IsNullOrEmpty(eurl.ReferenceAddress))
                {
                    var hyperlink = new HTMLElement(HtmlElements.A);
                    if (string.IsNullOrEmpty(eurl.AbsoluteUri))
                    {
                        hyperlink.AddAttribute("href", eurl.OriginalString);
                    }
                    else
                    {
                        hyperlink.AddAttribute("href", eurl.AbsoluteUri);
                    }
                    if (!string.IsNullOrEmpty(settings.HyperlinkTarget))
                    {
                        hyperlink.AddAttribute("target", settings.HyperlinkTarget);
                    }
                    hyperlink.Content = string.IsNullOrEmpty(cell.Text) ? eurl.Display : cell.Text;
                    element.AddChildElement(hyperlink);
                }
                else
                {
                    //Internal
                    element.Content = GetCellText(cell, settings);
                }
            }
            else
            {
                var hyperlink = new HTMLElement(HtmlElements.A);
                hyperlink.AddAttribute("href", cell.Hyperlink.OriginalString);
                if (!string.IsNullOrEmpty(settings.HyperlinkTarget))
                {
                    hyperlink.AddAttribute("target", settings.HyperlinkTarget);
                }
                hyperlink.Content = GetCellText(cell, settings);
                element.AddChildElement(hyperlink);
            }
        }

        protected void AddImage(HTMLElement parent, HtmlExportSettings settings, HtmlImage image, object value)
        {
            if (image != null)
            {
                var child = new HTMLElement(HtmlElements.Img);
                var name = GetPictureName(image);
                string imageName = HtmlExportTableUtil.GetClassName(image.Picture.Name, ((IPictureContainer)image.Picture).ImageHash);
                child.AddAttribute("alt", image.Picture.Name);
                if (settings.Pictures.AddNameAsId)
                {
                    child.AddAttribute("id", imageName);
                }
                child.AddAttribute("class", $"{settings.StyleClassPrefix}image-{name} {settings.StyleClassPrefix}image-prop-{imageName}");
                parent._childElements.Add(child);
            }
        }

        protected List<int> _columns = new List<int>();
        protected HtmlExportSettings Settings;
        protected readonly List<ExcelAddressBase> _mergedCells = new List<ExcelAddressBase>();

        protected void LoadVisibleColumns(ExcelRangeBase range)
        {
            var ws = range.Worksheet;
            _columns = new List<int>();
            for (int col = range._fromCol; col <= range._toCol; col++)
            {
                var c = ws.GetColumn(col);
                if (c == null || c.Hidden == false && c.Width > 0)
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

        protected void ValidateStream(Stream stream)
        {
            if (!stream.CanWrite)
            {
                throw new IOException("Parameter stream must be a writeable System.IO.Stream");
            }
        }

        internal bool HandleHiddenRow(EpplusHtmlAttribute attribute, ExcelWorksheet ws, HtmlExportSettings Settings, ref int row)
        {
            if (Settings.HiddenRows != eHiddenState.Include)
            {
                var r = ws.Row(row);
                if (r.Hidden || r.Height == 0)
                {
                    if (Settings.HiddenRows == eHiddenState.IncludeButHide)
                    {
                        attribute.AttributeName = "class";
                        attribute.Value = $"{Settings.StyleClassPrefix}hidden";
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

        internal void AddRowHeightStyle(HTMLElement element, ExcelRangeBase range, int row, string styleClassPrefix, bool isMultiSheet)
        {
            var r = range.Worksheet._values.GetValue(row, 0);
            if (r._value is RowInternal rowInternal)
            {
                if (rowInternal.Height != -1 && rowInternal.Height != range.Worksheet.DefaultRowHeight)
                {
                    element.AddAttribute("style", $"height:{rowInternal.Height}pt");
                    return;
                }
            }

            var clsName = HtmlExportTableUtil.GetWorksheetClassName(styleClassPrefix, "drh", range.Worksheet, isMultiSheet);
            element.AddAttribute("class", clsName); //Default row height
        }

        protected string GetPictureName(HtmlImage p)
        {
            var hash = ((IPictureContainer)p.Picture).ImageHash;
            var fi = new FileInfo(p.Picture.Part.Uri.OriginalString);
            var name = fi.Name.Substring(0, fi.Name.Length - fi.Extension.Length);

            return HtmlExportTableUtil.GetClassName(name, hash);
        }

        protected bool InMergeCellSpan(int row, int col)
        {
            for (int i = 0; i < _mergedCells.Count; i++)
            {
                var adr = _mergedCells[i];
                if (adr._toRow < row || adr._toRow == row && adr._toCol < col)
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

        protected void SetColRowSpan(ExcelRangeBase range, HTMLElement element, ExcelRange cell)
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
                            element.AddAttribute("colspan", colSpan.ToString(CultureInfo.InvariantCulture));
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
                            element.AddAttribute("rowspan", rowSpan.ToString(CultureInfo.InvariantCulture));
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

        protected void AddTableAccessibilityAttributes(AccessibilitySettings settings, HTMLElement element)
        {
            if (!settings.TableSettings.AddAccessibilityAttributes) return;
            if (!string.IsNullOrEmpty(settings.TableSettings.TableRole))
            {
                element.AddAttribute("role", settings.TableSettings.TableRole);
            }
            if (!string.IsNullOrEmpty(settings.TableSettings.AriaLabel))
            {
                element.AddAttribute(AriaAttributes.AriaLabel.AttributeName, settings.TableSettings.AriaLabel);
            }
            if (!string.IsNullOrEmpty(settings.TableSettings.AriaLabelledBy))
            {
                element.AddAttribute(AriaAttributes.AriaLabelledBy.AttributeName, settings.TableSettings.AriaLabelledBy);
            }
            if (!string.IsNullOrEmpty(settings.TableSettings.AriaDescribedBy))
            {
                element.AddAttribute(AriaAttributes.AriaDescribedBy.AttributeName, settings.TableSettings.AriaDescribedBy);
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

        protected void AddClassesAttributes(HTMLElement element, ExcelTable table, string tableId, List<string> additionalTableClassNames)
        {
            var tableClasses = TableClass;
            if (table != null)
            {
                tableClasses += " " + HtmlExportTableUtil.GetTableClasses(table); //Add classes for the table styles if the range corresponds to a table.
            }
            if (additionalTableClassNames != null && additionalTableClassNames.Count > 0)
            {
                foreach (var cls in additionalTableClassNames)
                {
                    tableClasses += $" {cls}";
                }
            }
            element.AddAttribute(HtmlAttributes.Class, $"{tableClasses}");

            if (!string.IsNullOrEmpty(tableId))
            {
                element.AddAttribute(HtmlAttributes.Id, tableId);
            }
        }

        internal void GetClassData(HTMLElement element, bool isTable, HtmlImage image, ExcelRangeBase cell, HtmlExportSettings settings, ExporterContext content, out HTMLElement valueElement, bool isHeader = false)
        {
            var imageCellClassName = GetImageCellClassName(image, Settings, isTable);
            var classString = AttributeTranslator.GetClassAttributeFromStyle(cell, isHeader, settings, imageCellClassName, content);
            var stylesAndExtras = AttributeTranslator.GetConditionalFormattings(cell, settings, content, ref classString);

            valueElement = element;

            if (!string.IsNullOrEmpty(classString))
            {
                element.AddAttribute("class", classString);
            }

            if (stylesAndExtras.Count > 0)
            {
                if (!string.IsNullOrEmpty(stylesAndExtras[0]) )
                {
                    element.AddAttribute("style", $"{stylesAndExtras[0]}");
                }
                if (stylesAndExtras.Count > 1)
                {
                    var childHtml = new HTMLElement("div");
                    childHtml.Content = stylesAndExtras[1];
                    element.AddChildElement(childHtml);
                }
            }
        }

        public void AddTableDataFromCell(ExcelRangeBase cell, string dataType, HTMLElement element, HtmlExportSettings settings, bool addRowScope, HtmlImage image, ExporterContext content)
        {
            if (dataType != ColumnDataTypeManager.HtmlDataTypes.String && settings.RenderDataAttributes)
            {
                var v = HtmlRawDataProvider.GetRawValue(cell.Value, dataType);
                if (string.IsNullOrEmpty(v) == false)
                {
                    element.AddAttribute($"data-{settings.DataValueAttributeName}", v);
                }
            }
            if (settings.Accessibility.TableSettings.AddAccessibilityAttributes)
            {
                element.AddAttribute("role", "cell");
                if (addRowScope)
                {
                    element.AddAttribute("scope", "row");
                }
            }

            GetClassData(element, true, image, cell, settings, content, out HTMLElement contentElement);

            AddImage(contentElement, settings, image, cell.Value);
            if (cell.IsRichText)
            {
                contentElement.Content = cell.RichText.HtmlText;
            }
            else
            {
                contentElement.Content = ValueToTextHandler.GetFormattedText(cell.Value, cell.Worksheet.Workbook, cell.StyleID, false, settings.Culture);
            }
        }
    }
}
