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
using OfficeOpenXml.Export.HtmlExport.Accessibility;
using OfficeOpenXml.Export.HtmlExport.HtmlCollections;
using OfficeOpenXml.Export.HtmlExport.Parsers;
using OfficeOpenXml.Table;
using OfficeOpenXml.Utils;
using System;
using System.Collections.Generic;
using System.IO;

namespace OfficeOpenXml.Export.HtmlExport.Exporters.Internal
{
    internal class HtmlRangeExporterSync : HtmlRangeExporterBase
    {
        internal HtmlRangeExporterSync
            (HtmlRangeExportSettings settings, ExcelRangeBase range) : base(settings, range)
        {
        }

        internal HtmlRangeExporterSync(HtmlRangeExportSettings settings, EPPlusReadOnlyList<ExcelRangeBase> ranges) : base(settings, ranges)
        {
        }

        /// <summary>
        /// Exports an <see cref="ExcelTable"/> to a html string
        /// </summary>
        /// <returns>A html table</returns>
        public string GetHtmlString()
        {
            using (var ms = RecyclableMemory.GetStream())
            {
                RenderHtml(ms, 0);
                ms.Position = 0;
                using (var sr = new StreamReader(ms))
                {
                    return sr.ReadToEnd();
                }
            }
        }
        /// <summary>
        /// Exports an <see cref="ExcelTable"/> to a html string
        /// </summary>
        /// <returns>A html table</returns>
        public string GetHtmlString(int rangeIndex)
        {
            ValidateRangeIndex(rangeIndex);
            using (var ms = RecyclableMemory.GetStream())
            {
                RenderHtml(ms, rangeIndex);
                ms.Position = 0;
                using (var sr = new StreamReader(ms))
                {
                    return sr.ReadToEnd();
                }
            }
        }

        /// <summary>
        /// Exports an <see cref="ExcelTable"/> to a html string
        /// </summary>
        /// <param name="rangeIndex">Index of the range to export</param>
        /// <param name="settings">Override some of the settings for this html exclusively</param>
        /// <returns>A html table</returns>
        public string GetHtmlString(int rangeIndex, ExcelHtmlOverrideExportSettings settings)
        {
            ValidateRangeIndex(rangeIndex);
            using (var ms = RecyclableMemory.GetStream())
            {
                RenderHtml(ms, rangeIndex, settings);
                ms.Position = 0;
                using (var sr = new StreamReader(ms))
                {
                    return sr.ReadToEnd();
                }
            }
        }

        /// <summary>
        /// Exports an <see cref="ExcelTable"/> to a html string
        /// </summary>
        /// <param name="rangeIndex">Index of the range to export</param>
        /// <param name="config">Override some of the settings for this html exclusively</param>
        /// <returns></returns>
        public string GetHtmlString(int rangeIndex, Action<ExcelHtmlOverrideExportSettings> config)
        {
            var settings = new ExcelHtmlOverrideExportSettings();
            config.Invoke(settings);
            return GetHtmlString(rangeIndex, settings);
        }

        /// <summary>
        /// Exports an <see cref="ExcelTable"/> to a html string
        /// </summary>
        /// <param name="stream">The stream to write to</param>
        /// <returns>A html table</returns>
        public void RenderHtml(Stream stream)
        {
            RenderHtml(stream, 0);
        }
        /// <summary>
        /// Exports an <see cref="ExcelTable"/> to a html string
        /// </summary>
        /// <param name="stream">The stream to write to</param>
        /// <param name="rangeIndex">The index of the range to output.</param>
        /// <param name="overrideSettings">Settings for this specific range index</param>
        /// <returns>A html table</returns>
        public void RenderHtml(Stream stream, int rangeIndex, ExcelHtmlOverrideExportSettings overrideSettings = null)
        {
            ValidateStream(stream);
            var htmlTable = GenerateHTML(rangeIndex, overrideSettings);

            var writer = new HtmlWriter(stream, Settings.Encoding);
            writer.RenderHTMLElement(htmlTable, Settings.Minify);

        }

        /// <summary>
        /// Exports an <see cref="ExcelTable"/> to a html string
        /// </summary>
        /// <param name="stream">The stream to write to</param>
        /// <param name="rangeIndex">Index of the range to export</param>
        /// <param name="config">Override some of the settings for this html exclusively</param>
        /// <returns></returns>
        public void RenderHtml(Stream stream, int rangeIndex, Action<ExcelHtmlOverrideExportSettings> config)
        {
            var settings = new ExcelHtmlOverrideExportSettings();
            config.Invoke(settings);
            RenderHtml(stream, rangeIndex, settings);
        }

        /// <summary>
        /// Renders both the Html and the Css to a single page. 
        /// </summary>
        /// <param name="htmlDocument">The html string where to insert the html and the css. The Html will be inserted in string parameter {0} and the Css will be inserted in parameter {1}.</param>
        /// <returns>The html document</returns>
        public string GetSinglePage(string htmlDocument = "<!DOCTYPE html>\r\n<html>\r\n<head>\r\n<style type=\"text/css\">\r\n{1}</style></head>\r\n<body>\r\n{0}\r\n</body>\r\n</html>")
        {
            if (Settings.Minify) htmlDocument = htmlDocument.Replace("\r\n", "");
            var html = GetHtmlString();
            var exporter = HtmlExporterFactory.CreateCssExporterSync(_settings, _ranges, _exporterContext);
            var css = exporter.GetCssString();
            return string.Format(htmlDocument, html, css);
        }
    }
}
