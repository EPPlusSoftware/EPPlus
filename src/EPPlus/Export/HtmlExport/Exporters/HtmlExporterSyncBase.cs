using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeOpenXml.Export.HtmlExport.Exporters
{
    internal abstract class HtmlExporterSyncBase : HtmlExporterBase
    {
        internal HtmlExporterSyncBase(ExcelRangeBase range) : base(range)
        {
        }

        internal HtmlExporterSyncBase(ExcelRangeBase[] ranges) : base(ranges)
        {
        }

        internal List<int> _columns = new List<int>();

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
    }
}
