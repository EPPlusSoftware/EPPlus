using OfficeOpenXml.Drawing.Interfaces;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Export.HtmlExport.Exporters
{
    internal abstract class HtmlExporterSyncBase : HtmlExporterBase
    {
        internal HtmlExporterSyncBase(HtmlExportSettings settings, ExcelRangeBase range) : base(settings, range)
        {
        }

        internal HtmlExporterSyncBase(HtmlExportSettings settings, ExcelRangeBase[] ranges) : base(settings, ranges)
        {
        }

        protected void SetColumnGroup(EpplusHtmlWriter writer, ExcelRangeBase _range, HtmlExportSettings settings, bool isMultiSheet)
        {
            var ws = _range.Worksheet;
            writer.RenderBeginTag("colgroup");
            writer.ApplyFormatIncreaseIndent(settings.Minify);
            var mdw = _range.Worksheet.Workbook.MaxFontWidth;
            var defColWidth = ExcelColumn.ColumnWidthToPixels(Convert.ToDecimal(ws.DefaultColWidth), mdw);
            foreach (var c in _columns)
            {
                if (settings.SetColumnWidth)
                {
                    double width = ws.GetColumnWidthPixels(c - 1, mdw);
                    if (width == defColWidth)
                    {
                        var clsName = GetWorksheetClassName(settings.StyleClassPrefix, "dcw", ws, isMultiSheet);
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
                writer.RenderBeginTag("col", true);
                writer.ApplyFormat(settings.Minify);
            }
            writer.Indent--;
            writer.RenderEndTag();
            writer.ApplyFormat(settings.Minify);
        }

        protected void AddImage(EpplusHtmlWriter writer, HtmlExportSettings settings, HtmlImage image, object value)
        {
            if (image != null)
            {
                var name = GetPictureName(image);
                string imageName = GetClassName(image.Picture.Name, ((IPictureContainer)image.Picture).ImageHash);
                writer.AddAttribute("alt", image.Picture.Name);
                if (settings.Pictures.AddNameAsId)
                {
                    writer.AddAttribute("id", imageName);
                }
                writer.AddAttribute("class", $"{settings.StyleClassPrefix}image-{name} {settings.StyleClassPrefix}image-prop-{imageName}");
                writer.RenderBeginTag("img", true);
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
    }
}
