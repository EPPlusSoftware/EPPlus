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
using OfficeOpenXml.Table;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
#if !NET35 && !NET40
using System.Threading.Tasks;

namespace OfficeOpenXml.Export.HtmlExport.Exporters
{
    internal abstract class HtmlRangeExporterAsyncBase : HtmlRangeExporterBase
    {
        internal HtmlRangeExporterAsyncBase
            (HtmlExportSettings settings, ExcelRangeBase range) : base(settings, range)
        {
            _settings = settings;
        }

        internal HtmlRangeExporterAsyncBase(HtmlExportSettings settings, EPPlusReadOnlyList<ExcelRangeBase> ranges) : base(settings, ranges)
        {
            _settings = settings;
        }

        private readonly HtmlExportSettings _settings;

        protected async Task RenderTableRowsAsync(ExcelRangeBase range, EpplusHtmlWriter writer, ExcelTable table, AccessibilitySettings accessibilitySettings, int headerRows)
        {
            if (accessibilitySettings.TableSettings.AddAccessibilityAttributes && !string.IsNullOrEmpty(accessibilitySettings.TableSettings.TbodyRole))
            {
                writer.AddAttribute("role", accessibilitySettings.TableSettings.TbodyRole);
            }
            await writer.RenderBeginTagAsync(HtmlElements.Tbody);
            await writer.ApplyFormatIncreaseIndentAsync(Settings.Minify);
            var row = range._fromRow + headerRows;
            var endRow = range._toRow;
            var ws = range.Worksheet;
            HtmlImage image = null;
            bool hasFooter = table != null && table.ShowTotal;
            while (row <= endRow)
            {
                if (HandleHiddenRow(writer, range.Worksheet, Settings, ref row))
                {
                    continue; //The row is hidden and should not be included.
                }

                if (hasFooter && row == endRow)
                {
                    await writer.RenderBeginTagAsync(HtmlElements.TFoot);
                }

                if (accessibilitySettings.TableSettings.AddAccessibilityAttributes)
                {
                    writer.AddAttribute("role", "row");
                    writer.AddAttribute("scope", "row");
                }

                if (Settings.SetRowHeight) AddRowHeightStyle(writer, range, row, Settings.StyleClassPrefix, IsMultiSheet);
                await writer.RenderBeginTagAsync(HtmlElements.TableRow);
                await writer.ApplyFormatIncreaseIndentAsync(Settings.Minify);
                foreach (var col in _columns)
                {
                    if (InMergeCellSpan(row, col)) continue;
                    var colIx = col - range._fromCol;
                    var cell = ws.Cells[row, col];
                    var cv = cell.Value;
                    var dataType = HtmlRawDataProvider.GetHtmlDataTypeFromValue(cell.Value);

                    SetColRowSpan(range, writer, cell);

                    if (Settings.Pictures.Include == ePictureInclude.Include)
                    {
                        image = GetImage(cell.Worksheet.PositionId, cell._fromRow, cell._fromCol);
                    }

                    if (cell.Hyperlink == null)
                    {
                        await _cellDataWriter.WriteAsync(cell, dataType, writer, Settings, accessibilitySettings, false, image);
                    }
                    else
                    {
                        var imageCellClassName = GetImageCellClassName(image, Settings);
                        writer.SetClassAttributeFromStyle(cell, false, Settings, imageCellClassName);
                        await writer.RenderBeginTagAsync(HtmlElements.TableData);
                        await AddImageAsync(writer, Settings, image, cell.Value);
                        await RenderHyperlinkAsync(writer, cell, Settings);
                        await writer.RenderEndTagAsync();
                        await writer.ApplyFormatAsync(Settings.Minify);
                    }
                }

                // end tag tr
                writer.Indent--;
                await writer.RenderEndTagAsync();
                await writer.ApplyFormatAsync(Settings.Minify);
                if (hasFooter && row == endRow)
                {
                    await writer.RenderEndTagAsync();
                }
                row++;
            }

            await writer.ApplyFormatDecreaseIndentAsync(Settings.Minify);
            // end tag tbody
            await writer.RenderEndTagAsync();
        }

        protected async Task RenderHeaderRowAsync(ExcelRangeBase range, EpplusHtmlWriter writer, ExcelTable table, AccessibilitySettings accessibilitySettings, int headerRows, List<string> headers)
        {
            if (table != null && table.ShowHeader == false) return;
            if (accessibilitySettings.TableSettings.AddAccessibilityAttributes && !string.IsNullOrEmpty(accessibilitySettings.TableSettings.TheadRole))
            {
                writer.AddAttribute("role", Settings.Accessibility.TableSettings.TheadRole);
            }
            await writer.RenderBeginTagAsync(HtmlElements.Thead);
            await writer.ApplyFormatIncreaseIndentAsync(Settings.Minify);
            if (table == null)
            {
                headerRows = headerRows == 0 ? 1 : headerRows;
            }
            else
            {
                headerRows = table.ShowHeader ? 1 : 0;
            }

            HtmlImage image = null;
            for (int i = 0; i < headerRows; i++)
            {
                if (accessibilitySettings.TableSettings.AddAccessibilityAttributes)
                {
                    writer.AddAttribute("role", "row");
                }
                var row = range._fromRow + i;
                if (Settings.SetRowHeight) AddRowHeightStyle(writer, range, row, Settings.StyleClassPrefix, IsMultiSheet);
                await writer.RenderBeginTagAsync(HtmlElements.TableRow);
                await writer.ApplyFormatIncreaseIndentAsync(Settings.Minify);
                foreach (var col in _columns)
                {
                    if (InMergeCellSpan(row, col)) continue;
                    var cell = range.Worksheet.Cells[row, col];
                    if (Settings.RenderDataTypes)
                    {
                        writer.AddAttribute("data-datatype", _dataTypes[col - range._fromCol]);
                    }
                    SetColRowSpan(range, writer, cell);
                    if (Settings.IncludeCssClassNames)
                    {
                        var imageCellClassName = GetImageCellClassName(image, Settings);
                        writer.SetClassAttributeFromStyle(cell, true, Settings, imageCellClassName);
                    }
                    if (Settings.Pictures.Include == ePictureInclude.Include)
                    {
                        image = GetImage(cell.Worksheet.PositionId, cell._fromRow, cell._fromCol);
                    }
                    await writer.RenderBeginTagAsync(HtmlElements.TableHeader);
                    await AddImageAsync(writer, Settings, image, cell.Value);

                    if (headerRows > 0 || table != null)
                    {
                        if (cell.Hyperlink == null)
                        {
                            await writer.WriteAsync(GetCellText(cell, Settings));
                        }
                        else
                        {
                            await RenderHyperlinkAsync(writer, cell, Settings);
                        }
                    }
                    else if (headers.Count < col)
                    {
                        writer.Write(headers[col]);
                    }

                    await writer.RenderEndTagAsync();
                    await writer.ApplyFormatAsync(Settings.Minify);
                }
                writer.Indent--;
                await writer.RenderEndTagAsync();
            }
            await writer.ApplyFormatDecreaseIndentAsync(Settings.Minify);
            await writer.RenderEndTagAsync();
            await writer.ApplyFormatAsync(Settings.Minify);
        }

        protected async Task RenderHyperlinkAsync(EpplusHtmlWriter writer, ExcelRangeBase cell, HtmlExportSettings settings)
        {
            if (cell.Hyperlink is ExcelHyperLink eurl)
            {
                if (string.IsNullOrEmpty(eurl.ReferenceAddress))
                {
                    if (string.IsNullOrEmpty(eurl.AbsoluteUri))
                    {
                        writer.AddAttribute("href", eurl.OriginalString);
                    }
                    else
                    {
                        writer.AddAttribute("href", eurl.AbsoluteUri);
                    }
                    if (!string.IsNullOrEmpty(settings.HyperlinkTarget))
                    {
                        writer.AddAttribute("target", settings.HyperlinkTarget);
                    }
                    await writer.RenderBeginTagAsync(HtmlElements.A);
                    await writer.WriteAsync(string.IsNullOrEmpty(eurl.Display) ? cell.Text : eurl.Display);
                    await writer.RenderEndTagAsync();
                }
                else
                {
                    //Internal
                    writer.Write(GetCellText(cell, settings));
                }
            }
            else
            {
                writer.AddAttribute("href", cell.Hyperlink.OriginalString);
                await writer.RenderBeginTagAsync(HtmlElements.A);
                await writer.WriteAsync(GetCellText(cell, settings));
                await writer.RenderEndTagAsync();
            }
        }

        protected async Task AddImageAsync(EpplusHtmlWriter writer, HtmlExportSettings settings, HtmlImage image, object value)
        {
            if (image != null)
            {
                var name = GetPictureName(image);
                string imageName = HtmlExportTableUtil.GetClassName(image.Picture.Name, ((IPictureContainer)image.Picture).ImageHash);
                writer.AddAttribute("alt", image.Picture.Name);
                if (settings.Pictures.AddNameAsId)
                {
                    writer.AddAttribute("id", imageName);
                }
                writer.AddAttribute("class", $"{settings.StyleClassPrefix}image-{name} {settings.StyleClassPrefix}image-prop-{imageName}");
                await writer.RenderBeginTagAsync(HtmlElements.Img, true);
            }
        }

        protected async Task SetColumnGroupAsync(EpplusHtmlWriter writer, ExcelRangeBase _range, HtmlExportSettings settings, bool isMultiSheet)
        {
            var ws = _range.Worksheet;
            await writer.RenderBeginTagAsync("colgroup");
            await writer.ApplyFormatIncreaseIndentAsync(settings.Minify);
            var mdw = _range.Worksheet.Workbook.MaxFontWidth;
            var defColWidth = ExcelColumn.ColumnWidthToPixels(Convert.ToDecimal(ws.DefaultColWidth), mdw);
            foreach (var c in _columns)
            {
                if (settings.SetColumnWidth)
                {
                    double width = ws.GetColumnWidthPixels(c - 1, mdw);
                    if (width == defColWidth)
                    {
                        var clsName = HtmlExportTableUtil.GetWorksheetClassName(settings.StyleClassPrefix, "dcw", ws, isMultiSheet);
                        writer.AddAttribute("class", clsName);
                    }
                    else
                    {
                        writer.AddAttribute("style", $"width:{width}px");
                    }
                }

                if (settings.HorizontalAlignmentWhenGeneral == eHtmlGeneralAlignmentHandling.ColumnDataType)
                {
                    writer.AddAttribute("class", $"{TableClass}-ar");
                }

                writer.AddAttribute("span", "1");
                await writer.RenderBeginTagAsync("col", true);
                await writer.ApplyFormatAsync(settings.Minify);
            }
            writer.Indent--;
            await writer.RenderEndTagAsync();
            await writer.ApplyFormatAsync(settings.Minify);
        }
    }
}
#endif
