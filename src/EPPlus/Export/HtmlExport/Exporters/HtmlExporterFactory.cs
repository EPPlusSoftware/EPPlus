using OfficeOpenXml.Table;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Export.HtmlExport.Exporters
{
    internal static class HtmlExporterFactory
    {
        public static HtmlExporterSync CreateHtmlExporterSync(HtmlRangeExportSettings settings, ExcelRangeBase range)
        {
            return new HtmlExporterSync(settings, range);
        }

        public static HtmlExporterSync CreateHtmlExporterSync(HtmlRangeExportSettings settings, ExcelRangeBase[] ranges)
        {
            return new HtmlExporterSync(settings, ranges);
        }

        public static HtmlTableExporterSync CreateHtmlTableExporterSync(HtmlTableExportSettings settings, ExcelTable table)
        {
            return new HtmlTableExporterSync(settings, table);
        }

        public static CssExporterSync CreateCssExporterSync(HtmlRangeExportSettings settings, ExcelRangeBase range)
        {
            return new CssExporterSync(settings, range);
        }

        public static CssExporterSync CreateCssExporterSync(HtmlRangeExportSettings settings, ExcelRangeBase[] ranges)
        {
            return new CssExporterSync(settings, ranges);
        }

        public static CssExporterTableSync CreateCssExporterTableSync(HtmlTableExportSettings settings, ExcelTable table)
        {
            return new CssExporterTableSync(settings, table);
        }

#if !NET35 && !NET40
        public static HtmlExporterAsync CreateHtmlExporterAsync(HtmlRangeExportSettings settings, ExcelRangeBase range)
        {
            return new HtmlExporterAsync(settings, range);
        }

        public static HtmlExporterAsync CreateHtmlExporterAsync(HtmlRangeExportSettings settings, ExcelRangeBase[] ranges)
        {
            return new HtmlExporterAsync(settings, ranges);
        }

        public static HtmlTableExporterAsync CreateHtmlTableExporterAsync(HtmlTableExportSettings settings, ExcelTable table)
        {
            return new HtmlTableExporterAsync(settings, table);
        }

        public static CssExporterAsync CreateCssExporterAsync(HtmlRangeExportSettings settings, ExcelRangeBase range)
        {
            return new CssExporterAsync(settings, range);
        }

        public static CssExporterAsync CreateCssExporterAsync(HtmlRangeExportSettings settings, ExcelRangeBase[] ranges)
        {
            return new CssExporterAsync(settings, ranges);
        }

        public static CssExporterTableAsync CreateCssExporterTableAsync(HtmlTableExportSettings settings, ExcelTable table)
        {
            return new CssExporterTableAsync(settings, table);
        }
#endif
    }
}

