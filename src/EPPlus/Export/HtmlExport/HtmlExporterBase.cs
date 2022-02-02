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
using System;
using System.Collections.Generic;

namespace OfficeOpenXml.Export.HtmlExport
{
    /// <summary>
    /// Baseclass for html exporters
    /// </summary>
    public abstract partial class HtmlExporterBase
    {
        internal const string TableClass = "epplus-table";
        internal List<string> _datatypes = new List<string>();
        internal List<int> _columns = new List<int>();

        internal void AddRowHeightStyle(EpplusHtmlWriter writer, ExcelRangeBase range, int row, string styleClassPrefix)
        {
            var r = range.Worksheet._values.GetValue(row, 0);
            if (r._value is RowInternal rowInternal)
            {
                if (rowInternal.Height != range.Worksheet.DefaultRowHeight)
                {
                    var ht = range.Worksheet.GetRowHeight(row);
                    writer.AddAttribute("style", $"height:{ht}px");
                }
            }
            writer.AddAttribute("class", $"{styleClassPrefix}drh"); //Default row height
        }
        internal void SetColumnGroup(EpplusHtmlWriter writer, ExcelRangeBase _range, HtmlExportSettings settings)
        {
            var ws = _range.Worksheet;
            writer.RenderBeginTag("colgroup");
            writer.Indent++;
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
                writer.RenderBeginTag("col", true);

            }
            writer.Indent--;
            writer.RenderEndTag();
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

        internal static string GetClassName(string className)
        {
            className = className.Trim().Replace(" ", "-");
            var newClassName = "";
            for (int i = 0; i < className.Length; i++)
            {
                var c = className[i];
                if (i == 0)
                {
                    if (c == '-' || (c >= '0' && c <= '9'))
                    {
                        newClassName = "_";
                    }
                }

                if ((c >= '0' && c <= '9') ||
                   (c >= 'a' && c <= 'z') ||
                   (c >= 'A' && c <= 'Z') ||
                    c >= 0x00A0)
                {
                    newClassName += c;
                }
            }
            return string.IsNullOrEmpty(newClassName) ? "EmptyClass" : newClassName;
        }
    }
}