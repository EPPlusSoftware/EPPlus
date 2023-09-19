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
using OfficeOpenXml.Export.HtmlExport.Settings;
using OfficeOpenXml.Table;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Export.HtmlExport.Exporters
{
    internal static class HtmlExporterFactory
    {

        public static HtmlRangeExporterSync CreateHtmlExporterSync(HtmlRangeExportSettings settings, ExcelRangeBase range, Dictionary<string, int> styleCache)
        {
            var exporter = new HtmlRangeExporterSync(settings, range);
            exporter.SetStyleCache(styleCache);
            return exporter;
        }

        public static HtmlRangeExporterSync CreateHtmlExporterSync(HtmlRangeExportSettings settings, EPPlusReadOnlyList<ExcelRangeBase> ranges, Dictionary<string, int> styleCache)
        {
            var exporter = new HtmlRangeExporterSync(settings, ranges);
            exporter.SetStyleCache(styleCache);
            return exporter;
        }

        public static HtmlTableExporterSync CreateHtmlTableExporterSync(HtmlTableExportSettings settings, ExcelTable table, Dictionary<string, int> styleCache)
        {
            var exporter = new HtmlTableExporterSync(settings, table);
            exporter.SetStyleCache(styleCache);
            return exporter;
        }

        public static CssRangeExporterSync CreateCssExporterSync(HtmlRangeExportSettings settings, ExcelRangeBase range, Dictionary<string, int> styleCache)
        {
            var exporter = new CssRangeExporterSync(settings, range);
            exporter.SetStyleCache(styleCache);
            return exporter;
        }

        public static CssRangeExporterSync CreateCssExporterSync(HtmlRangeExportSettings settings, EPPlusReadOnlyList<ExcelRangeBase> ranges, Dictionary<string, int> styleCache)
        {
            var exporter = new CssRangeExporterSync(settings, ranges);
            exporter.SetStyleCache(styleCache);
            return exporter;
        }

        public static CssTableExporterSync CreateCssExporterTableSync(HtmlTableExportSettings settings, ExcelTable table, Dictionary<string, int> styleCache)
        {
            var exporter = new CssTableExporterSync(settings, table);
            exporter.SetStyleCache(styleCache);
            return exporter;
        }

#if !NET35 && !NET40
        public static HtmlRangeExporterAsync CreateHtmlExporterAsync(HtmlRangeExportSettings settings, ExcelRangeBase range, Dictionary<string, int> styleCache)
        {
            var exporter = new HtmlRangeExporterAsync(settings, range);
            exporter.SetStyleCache(styleCache);
            return exporter;
        }

        public static HtmlRangeExporterAsync CreateHtmlExporterAsync(HtmlRangeExportSettings settings, EPPlusReadOnlyList<ExcelRangeBase> ranges, Dictionary<string, int> styleCache)
        {
            var exporter = new HtmlRangeExporterAsync(settings, ranges);
            exporter.SetStyleCache(styleCache);
            return exporter;
        }

        public static HtmlTableExporterAsync CreateHtmlTableExporterAsync(HtmlTableExportSettings settings, ExcelTable table, Dictionary<string, int> styleCache)
        {
            var exporter = new HtmlTableExporterAsync(settings, table);
            exporter.SetStyleCache(styleCache);
            return exporter;
        }

        public static CssRangeExporterAsync CreateCssExporterAsync(HtmlRangeExportSettings settings, ExcelRangeBase range, Dictionary<string, int> styleCache)
        {
            var exporter = new CssRangeExporterAsync(settings, range);
            exporter.SetStyleCache(styleCache);
            return exporter;
        }

        public static CssRangeExporterAsync CreateCssExporterAsync(HtmlRangeExportSettings settings, EPPlusReadOnlyList<ExcelRangeBase> ranges, Dictionary<string, int> styleCache)
        {
            var exporter = new CssRangeExporterAsync(settings, ranges);
            exporter.SetStyleCache(styleCache);
            return exporter;
        }

        public static CssTableExporterAsync CreateCssExporterTableAsync(HtmlTableExportSettings settings, ExcelTable table, Dictionary<string, int> styleCache)
        {
            var exporter = new CssTableExporterAsync(settings, table);
            exporter.SetStyleCache(styleCache);
            return exporter;
        }
#endif
    }
}

