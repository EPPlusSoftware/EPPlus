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

namespace OfficeOpenXml.Export.HtmlExport.Exporters
{
    internal abstract class HtmlRangeExporterSyncBase : HtmlRangeExporterBase
    {
        internal HtmlRangeExporterSyncBase(HtmlExportSettings settings, ExcelRangeBase range) : base(settings, range)
        {
        }

        internal HtmlRangeExporterSyncBase(HtmlExportSettings settings, EPPlusReadOnlyList<ExcelRangeBase> ranges) : base(settings, ranges)
        {
        }

        protected void SetColumnGroup(HTMLElement element, ExcelRangeBase _range, HtmlExportSettings settings, bool isMultiSheet)
        {
            var group = GetGroup(_range, settings, isMultiSheet);
            element.AddChildElement(group);
            //writer.RenderHTMLElement(group, settings.Minify);
        }

        HTMLElement GetGroup(ExcelRangeBase _range, HtmlExportSettings settings, bool isMultiSheet)
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
                        // writer.AddAttribute("class", clsName);
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
                //parent.RenderBeginTag(HtmlElements.Img, true);
            }
        }

        protected void AddImage(EpplusHtmlWriter writer, HtmlExportSettings settings, HtmlImage image, object value)
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
                writer.RenderBeginTag(HtmlElements.Img, true);
            }
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
                    if (string.IsNullOrEmpty(eurl.AbsoluteUri))
                    {
                        element.AddAttribute("href", eurl.OriginalString);
                    }
                    else
                    {
                        element.AddAttribute("href", eurl.AbsoluteUri);
                    }
                    if (!string.IsNullOrEmpty(settings.HyperlinkTarget))
                    {
                        element.AddAttribute("target", settings.HyperlinkTarget);
                    }
                    var hyperlink = new HTMLElement(HtmlElements.A);
                    hyperlink.Content = string.IsNullOrEmpty(eurl.Display) ? cell.Text : eurl.Display;
                    element.AddChildElement(hyperlink);
                    //writer.RenderBeginTag(HtmlElements.A);
                    //writer.Write(string.IsNullOrEmpty(eurl.Display) ? cell.Text : eurl.Display);
                    //writer.RenderEndTag();
                }
                else
                {
                    //Internal
                    element.Content = GetCellText(cell, settings);
                   // writer.Write(GetCellText(cell, settings));
                }
            }
            else
            {
                element.AddAttribute("href", cell.Hyperlink.OriginalString);
                //writer.AddAttribute("href", cell.Hyperlink.OriginalString);
                if (!string.IsNullOrEmpty(settings.HyperlinkTarget))
                {
                    element.AddAttribute("target", settings.HyperlinkTarget);
                }
                var hyperlink = new HTMLElement(HtmlElements.A);
                hyperlink.Content = GetCellText(cell, settings);
                element.AddChildElement(hyperlink);
                //writer.RenderBeginTag(HtmlElements.A);
                //writer.Write(GetCellText(cell, settings));
                //writer.RenderEndTag();
            }
        }

        /// <summary>
        /// Renders a hyperlink
        /// </summary>
        /// <param name="writer"></param>
        /// <param name="cell"></param>
        /// <param name="settings"></param>
        protected void RenderHyperlink(EpplusHtmlWriter writer, ExcelRangeBase cell, HtmlExportSettings settings)
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
                    writer.RenderBeginTag(HtmlElements.A);
                    writer.Write(string.IsNullOrEmpty(eurl.Display) ? cell.Text : eurl.Display);
                    writer.RenderEndTag();
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
                if (!string.IsNullOrEmpty(settings.HyperlinkTarget))
                {
                    writer.AddAttribute("target", settings.HyperlinkTarget);
                }
                writer.RenderBeginTag(HtmlElements.A);
                writer.Write(GetCellText(cell, settings));
                writer.RenderEndTag();
            }
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

        protected HTMLElement AddTableRowsAlt(ExcelRangeBase range, int row, int endRow)
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

                    if (!table.ShowFirstColumn && !table.ShowLastColumn || table == null)
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
                    var cv = cell.Value;
                    var dataType = HtmlRawDataProvider.GetHtmlDataTypeFromValue(cell.Value);

                    var tblData = new HTMLElement(HtmlElements.TableData);

                    SetColRowSpan(range, tblData, cell);

                    if (Settings.Pictures.Include == ePictureInclude.Include)
                    {
                        image = GetImage(cell.Worksheet.PositionId, cell._fromRow, cell._fromCol);
                    }

                    if (cell.Hyperlink == null)
                    {
                        _cellDataWriter.Write(cell, dataType, tblData, Settings, accessibilitySettings, false, image, _exporterContext);
                    }
                    else
                    {
                        var imageCellClassName = GetImageCellClassName(image, Settings);

                        var classString = AttributeTranslator.GetClassAttributeFromStyle(cell, false, Settings, imageCellClassName, _exporterContext);

                        if (!string.IsNullOrEmpty(classString))
                        {
                            tblData.AddAttribute("class", classString);
                        }

                        AddImage(tblData, Settings, image, cell.Value);
                        AddHyperlink(tblData, cell, Settings);
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

            element.AddChildElement(tBody);
        }

        protected HTMLElement GetTheadAlt(ExcelRangeBase range, List<string> headers = null)
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
                    if(InMergeCellSpan(row, col)) continue;

                    var th = new HTMLElement(HtmlElements.TableHeader);
                    var cell = worksheet.Cells[row, col];
                    if (Settings.RenderDataTypes)
                    {
                        th.AddAttribute("data-datatype", _dataTypes[col - range._fromCol]);
                    }

                    SetColRowSpan(range, th, cell);

                    if (Settings.IncludeCssClassNames)
                    {
                        var imageCellClassName = GetImageCellClassName(image, Settings, table != null);
                        var classString = AttributeTranslator.GetClassAttributeFromStyle(cell, true, Settings, imageCellClassName, _exporterContext);

                        if (!string.IsNullOrEmpty(classString))
                        {
                            th.AddAttribute("class", classString);
                        }
                    }

                    AddTableData(table, th, col);

                    if (Settings.Pictures.Include == ePictureInclude.Include)
                    {
                        image = GetImage(cell.Worksheet.PositionId, cell._fromRow, cell._fromCol);
                    }
                    AddImage(th, Settings, image, cell.Value);

                    if (headerRows > 0 || table != null)
                    {
                        if (cell.Hyperlink == null)
                        {
                            th.Content = GetCellText(cell, Settings);
                        }
                        else
                        {
                            AddHyperlink(th, cell, Settings);
                        }
                    }
                    else if (headers.Count < col)
                    {
                        th.Content = headers[col];
                    }

                    rowElement.AddChildElement(th);
                }
                thead.AddChildElement(rowElement);
            }
            return thead;
        }
    }
}
