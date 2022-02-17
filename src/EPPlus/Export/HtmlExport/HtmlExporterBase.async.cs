/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  01/17/2022         EPPlus Software AB           ExcelTable Html Export
 *************************************************************************************************/
#if !NET35 && !NET40
using System;
using System.Threading.Tasks;
namespace OfficeOpenXml.Export.HtmlExport
{
    /// <summary>
    /// Baseclass for html exporters
    /// </summary>
    public abstract partial class HtmlExporterBase
    {
        internal async Task SetColumnGroupAsync(EpplusHtmlWriter writer, ExcelRangeBase _range, HtmlExportSettings settings)
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
                        writer.AddAttribute("class", $"{settings.StyleClassPrefix}dcw");
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
