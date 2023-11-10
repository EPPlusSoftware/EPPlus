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
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

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

        protected void SetColumnGroup(EpplusHtmlWriter writer, ExcelRangeBase _range, HtmlExportSettings settings, bool isMultiSheet)
        {
            //writer.RenderBeginTag("colgroup");
            //writer.ApplyFormatIncreaseIndent(settings.Minify);
            var group = new HTMLElement("colgroup");

            var ws = _range.Worksheet;
            var mdw = _range.Worksheet.Workbook.MaxFontWidth;
            var defColWidth = ExcelColumn.ColumnWidthToPixels(Convert.ToDecimal(ws.DefaultColWidth), mdw);

            //string classes = "";
            //string args = "";
            //var elements = new List<HTMLElement>();

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
                //writer.RenderBeginTag("col", true);
                //writer.ApplyFormat(settings.Minify);
            }

            //foreach (var c in _columns)
            //{
            //    if (settings.SetColumnWidth)
            //    {
            //        double width = ws.GetColumnWidthPixels(c - 1, mdw);
            //        if (width == defColWidth)
            //        {
            //            var clsName = HtmlExportTableUtil.GetWorksheetClassName(settings.StyleClassPrefix, "dcw", ws, isMultiSheet);
            //            writer.AddAttribute("class", clsName);
            //        }
            //        else
            //        {
            //            writer.AddAttribute("style", $"width:{width}px");
            //        }
            //    }
            //    if (settings.HorizontalAlignmentWhenGeneral == eHtmlGeneralAlignmentHandling.ColumnDataType)
            //    {
            //        writer.AddAttribute("class", $"{TableClass}-ar");
            //    }
            //    writer.AddAttribute("span", "1");
            //    writer.RenderBeginTag("col", true);
            //    writer.ApplyFormat(settings.Minify);
            //}

            //writer.Indent--;
            //writer.RenderEndTag();
            //writer.ApplyFormat(settings.Minify);
            writer.RenderHTMLElement(group, settings.Minify);
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
