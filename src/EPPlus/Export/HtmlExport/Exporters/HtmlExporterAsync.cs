using OfficeOpenXml.Table;
using OfficeOpenXml.Utils;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
#if !NET35 && !NET40
using System.Threading.Tasks;

namespace OfficeOpenXml.Export.HtmlExport.Exporters
{
    internal class HtmlExporterAsync : HtmlExporterAsyncBase
    {
        internal HtmlExporterAsync
           (HtmlRangeExportSettings settings, ExcelRangeBase range) : base(settings, range)
        {
            _settings = settings;
        }

        internal HtmlExporterAsync(HtmlRangeExportSettings settings, ExcelRangeBase[] ranges) : base(settings, ranges)
        {
            _settings = settings;
        }

        private readonly HtmlRangeExportSettings _settings;

        /// <summary>
        /// Exports an <see cref="ExcelTable"/> to a html string
        /// </summary>
        /// <returns>A html table</returns>
        public async Task<string> GetHtmlStringAsync()
        {
            using (var ms = RecyclableMemory.GetStream())
            {
                await RenderHtmlAsync(ms, 0);
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
        public async Task<string> GetHtmlStringAsync(int rangeIndex, ExcelHtmlOverrideExportSettings settings = null)
        {
            using (var ms = RecyclableMemory.GetStream())
            {
                await RenderHtmlAsync(ms, rangeIndex, settings);
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
        public async Task<string> GetHtmlStringAsync(int rangeIndex, Action<ExcelHtmlOverrideExportSettings> config)
        {
            var settings = new ExcelHtmlOverrideExportSettings();
            config.Invoke(settings);
            return await GetHtmlStringAsync(rangeIndex, settings);
        }

        /// <summary>
        /// Exports an <see cref="ExcelTable"/> to a html string
        /// </summary>
        /// <param name="stream">The stream to write to</param>
        /// <returns>A html table</returns>
        public async Task RenderHtmlAsync(Stream stream)
        {
            await RenderHtmlAsync(stream, 0);
        }

        /// <summary>
        /// Exports the html part of the html export, without the styles.
        /// </summary>
        /// <param name="stream">The stream to write to.</param>
        /// <param name="rangeIndex">The index of the range to output.</param>
        /// <param name="overrideSettings">Settings for this specific range index</param>
        /// <exception cref="IOException"></exception>
        public async Task RenderHtmlAsync(Stream stream, int rangeIndex, ExcelHtmlOverrideExportSettings overrideSettings = null)
        {
            ValidateRangeIndex(rangeIndex);
            if (!stream.CanWrite)
            {
                throw new IOException("Parameter stream must be a writeable System.IO.Stream");
            }
            _mergedCells.Clear();
            var range = _ranges[rangeIndex];
            GetDataTypes(_ranges[rangeIndex], _settings);


            ExcelTable table = null;
            if (Settings.TableStyle != eHtmlRangeTableInclude.Exclude)
            {
                table = range.GetTable();
            }
            var writer = new EpplusHtmlWriter(stream, Settings.Encoding, _styleCache);
            var tableId = GetTableId(rangeIndex, overrideSettings);
            var additionalClassNames = GetAdditionalClassNames(overrideSettings);
            var accessibilitySettings = GetAccessibilitySettings(overrideSettings);
            var headerRows = overrideSettings != null ? overrideSettings.HeaderRows : _settings.HeaderRows;
            var headers = overrideSettings != null ? overrideSettings.Headers : _settings.Headers;
            AddClassesAttributes(writer, table, tableId, additionalClassNames);
            AddTableAccessibilityAttributes(accessibilitySettings, writer);
            await writer.RenderBeginTagAsync(HtmlElements.Table);

            await writer.ApplyFormatIncreaseIndentAsync(Settings.Minify);
            LoadVisibleColumns(range);
            if (Settings.SetColumnWidth || Settings.HorizontalAlignmentWhenGeneral == eHtmlGeneralAlignmentHandling.ColumnDataType)
            {
                await SetColumnGroupAsync(writer, range, Settings, IsMultiSheet);
            }

            if (headerRows > 0 || headers.Count > 0)
            {
                await RenderHeaderRowAsync(range, writer, table, headerRows, headers);
            }
            // table rows
            await RenderTableRowsAsync(range, writer, table, accessibilitySettings, _settings.HeaderRows);

            // end tag table
            await writer.RenderEndTagAsync();
        }

        /// <summary>
        /// Exports the html part of the html export, without the styles.
        /// </summary>
        /// <param name="stream">The stream to write to.</param>
        /// <param name="rangeIndex">Index of the range to export</param>
        /// <param name="config">Override some of the settings for this html exclusively</param>
        /// <returns></returns>
        public async Task RenderHtmlAsync(Stream stream, int rangeIndex, Action<ExcelHtmlOverrideExportSettings> config)
        {
            var settings = new ExcelHtmlOverrideExportSettings();
            config.Invoke(settings);
            await RenderHtmlAsync(stream, rangeIndex, settings);
        }

        /// <summary>
        /// Renders the first range of the Html and the Css to a single page. 
        /// </summary>
        /// <param name="htmlDocument">The html string where to insert the html and the css. The Html will be inserted in string parameter {0} and the Css will be inserted in parameter {1}.</param>
        /// <returns>The html document</returns>
        public async Task<string> GetSinglePageAsync(string htmlDocument = "<!DOCTYPE html>\r\n<html>\r\n<head>\r\n<style type=\"text/css\">\r\n{1}</style></head>\r\n<body>\r\n{0}</body>\r\n</html>")
        {
            if (Settings.Minify) htmlDocument = htmlDocument.Replace("\r\n", "");
            var html = await GetHtmlStringAsync();
            var cssExporter = HtmlExporterFactory.CreateCssExporterAsync(_settings, _ranges.ToArray());
            var css = await cssExporter.GetCssStringAsync();
            return string.Format(htmlDocument, html, css);
        }
    }
}
#endif
