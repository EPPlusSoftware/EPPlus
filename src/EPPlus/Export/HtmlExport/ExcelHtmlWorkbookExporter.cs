using OfficeOpenXml.Export.HtmlExport.Exporters;
using OfficeOpenXml.Export.HtmlExport.Interfaces;
using OfficeOpenXml.Table;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
#if !NET35 && !NET40
using System.Threading.Tasks;
#endif

namespace OfficeOpenXml.Export.HtmlExport
{
    internal class ExcelHtmlWorkbookExporter : IExcelHtmlRangeExporter
    {
        public ExcelHtmlWorkbookExporter(params ExcelRangeBase[] ranges)
        {
            _ranges = ranges;
            _settings = new HtmlRangeExportSettings();
        }

        private readonly ExcelRangeBase[] _ranges;
        private readonly HtmlRangeExportSettings _settings;

        public HtmlRangeExportSettings Settings
            { get { return _settings; } }   

        /// <summary>
        /// Exports an <see cref="ExcelTable"/> to a html string
        /// </summary>
        /// <returns>A html table</returns>
        public string GetHtmlString()
        {
            var exporter = HtmlExporterFactory.CreateHtmlExporterSync(_settings, _ranges);
            return exporter.GetHtmlString();
        }

        /// <summary>
        /// Exports an <see cref="ExcelTable"/> to a html string
        /// </summary>
        /// <param name="rangeIndex">0-based index of the requested range</param>
        /// <returns>A html table</returns>
        public string GetHtmlString(int rangeIndex)
        {
            var exporter = HtmlExporterFactory.CreateHtmlExporterSync(_settings, _ranges);
            return exporter.GetHtmlString(rangeIndex);
        }

        /// <summary>
        /// Exports an <see cref="ExcelTable"/> to a html string
        /// </summary>
        /// <param name="rangeIndex">Index of the range to export</param>
        /// <param name="settings">Override some of the settings for this html exclusively</param>
        /// <returns>A html table</returns>
        public string GetHtmlString(int rangeIndex, ExcelHtmlOverrideExportSettings settings)
        {
            var exporter = HtmlExporterFactory.CreateHtmlExporterSync(_settings, _ranges);
            return exporter.GetHtmlString(rangeIndex, settings);
        }

        /// <summary>
        /// Exports an <see cref="ExcelTable"/> to a html string
        /// </summary>
        /// <param name="rangeIndex">Index of the range to export</param>
        /// <param name="config">Override some of the settings for this html exclusively</param>
        /// <returns></returns>
        public string GetHtmlString(int rangeIndex, Action<ExcelHtmlOverrideExportSettings> config)
        {
            var exporter = HtmlExporterFactory.CreateHtmlExporterSync(_settings, _ranges);
            return exporter.GetHtmlString(rangeIndex, config);
        }

        /// <summary>
        /// Exports an <see cref="ExcelTable"/> to a html string
        /// </summary>
        /// <param name="stream">The stream to write to</param>
        /// <returns>A html table</returns>
        public void RenderHtml(Stream stream)
        {
            var exporter = HtmlExporterFactory.CreateHtmlExporterSync(_settings, _ranges);
            exporter.RenderHtml(stream);
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
            var exporter = HtmlExporterFactory.CreateHtmlExporterSync(_settings, _ranges);
            exporter.RenderHtml(stream, rangeIndex, overrideSettings);
        }

        /// <summary>
        /// Exports an <see cref="ExcelTable"/> to a html string
        /// </summary>
        /// <param name="stream">The stream to write to</param>
        /// <param name="rangeIndex">The index of the range to output.</param>
        /// <param name="config">Settings for this specific range index</param>
        /// <returns>A html table</returns>
        public void RenderHtml(Stream stream, int rangeIndex, Action<ExcelHtmlOverrideExportSettings> config)
        {
            var exporter = HtmlExporterFactory.CreateHtmlExporterSync(_settings, _ranges);
            exporter.RenderHtml(stream, rangeIndex, config);
        }

        /// <summary>
        /// Renders both the Html and the Css to a single page. 
        /// </summary>
        /// <param name="htmlDocument">The html string where to insert the html and the css. The Html will be inserted in string parameter {0} and the Css will be inserted in parameter {1}.</param>
        /// <returns>The html document</returns>
        public string GetSinglePage(string htmlDocument = "<!DOCTYPE html>\r\n<html>\r\n<head>\r\n<style type=\"text/css\">\r\n{1}</style></head>\r\n<body>\r\n{0}</body>\r\n</html>")
        {
            var exporter = HtmlExporterFactory.CreateHtmlExporterSync(_settings, _ranges);
            return exporter.GetSinglePage(htmlDocument);
        }

        /// <summary>
        /// Exports an <see cref="ExcelTable"/> to a html string
        /// </summary>
        /// <returns>Cascading style sheet for the exported range</returns>
        public string GetCssString()
        {
            var exporter = HtmlExporterFactory.CreateCssExporterSync(_settings, _ranges);
            return exporter.GetCssString();
        }

        /// <summary>
        /// Exports the css part of the html export.
        /// </summary>
        /// <param name="stream">The stream to write the css to.</param>
        /// <exception cref="IOException"></exception>
        public void RenderCss(Stream stream)
        {
            var exporter = HtmlExporterFactory.CreateCssExporterSync(_settings, _ranges);
            exporter.RenderCss(stream);
        }

#if !NET35 && !NET40
        /// <summary>
        /// Exports an <see cref="ExcelTable"/> to a html string
        /// </summary>
        /// <returns>A html table</returns>
        public Task<string> GetHtmlStringAsync()
        {
            var exporter = HtmlExporterFactory.CreateHtmlExporterAsync(_settings, _ranges);
            return exporter.GetHtmlStringAsync();
        }

        /// <summary>
        /// Exports an <see cref="ExcelTable"/> to a html string
        /// </summary>
        /// <param name="rangeIndex">Index of the range to export</param>
        /// <param name="settings">Override some of the settings for this html exclusively</param>
        /// <returns>A html table</returns>
        public Task<string> GetHtmlStringAsync(int rangeIndex, ExcelHtmlOverrideExportSettings settings = null)
        {
            var exporter = HtmlExporterFactory.CreateHtmlExporterAsync(_settings, _ranges);
            return exporter.GetHtmlStringAsync(rangeIndex, settings);
        }

        /// <summary>
        /// Exports an <see cref="ExcelTable"/> to a html string
        /// </summary>
        /// <param name="rangeIndex">Index of the range to export</param>
        /// <param name="config">Override some of the settings for this html exclusively</param>
        /// <returns></returns>
        public Task<string> GetHtmlStringAsync(int rangeIndex, Action<ExcelHtmlOverrideExportSettings> config)
        {
            var exporter = HtmlExporterFactory.CreateHtmlExporterAsync(_settings, _ranges);
            return exporter.GetHtmlStringAsync(rangeIndex, config);
        }

        /// <summary>
        /// Exports an <see cref="ExcelTable"/> to a html string
        /// </summary>
        /// <param name="stream">The stream to write to</param>
        /// <returns>A html table</returns>
        public Task RenderHtmlAsync(Stream stream)
        {
            var exporter = HtmlExporterFactory.CreateHtmlExporterAsync(_settings, _ranges);
            return exporter.RenderHtmlAsync(stream);
        }

        /// <summary>
        /// Exports the html part of the html export, without the styles.
        /// </summary>
        /// <param name="stream">The stream to write to.</param>
        /// <param name="rangeIndex">The index of the range to output.</param>
        /// <param name="overrideSettings">Settings for this specific range index</param>
        /// <exception cref="IOException"></exception>
        public Task RenderHtmlAsync(Stream stream, int rangeIndex, ExcelHtmlOverrideExportSettings overrideSettings = null)
        {
            var exporter = HtmlExporterFactory.CreateHtmlExporterAsync(_settings, _ranges);
            return exporter.RenderHtmlAsync(stream, rangeIndex, overrideSettings);
        }

        /// <summary>
        /// Exports the html part of the html export, without the styles.
        /// </summary>
        /// <param name="stream">The stream to write to.</param>
        /// <param name="rangeIndex">Index of the range to export</param>
        /// <param name="config">Override some of the settings for this html exclusively</param>
        /// <returns></returns>
        public Task RenderHtmlAsync(Stream stream, int rangeIndex, Action<ExcelHtmlOverrideExportSettings> config)
        {
            var exporter = HtmlExporterFactory.CreateHtmlExporterAsync(_settings, _ranges);
            return exporter.RenderHtmlAsync(stream, rangeIndex, config);
        }

        /// <summary>
        /// Renders the first range of the Html and the Css to a single page. 
        /// </summary>
        /// <param name="htmlDocument">The html string where to insert the html and the css. The Html will be inserted in string parameter {0} and the Css will be inserted in parameter {1}.</param>
        /// <returns>The html document</returns>
        public Task<string> GetSinglePageAsync(string htmlDocument = "<!DOCTYPE html>\r\n<html>\r\n<head>\r\n<style type=\"text/css\">\r\n{1}</style></head>\r\n<body>\r\n{0}</body>\r\n</html>")
        {
            var exporter = HtmlExporterFactory.CreateHtmlExporterAsync(_settings, _ranges);
            return exporter.GetSinglePageAsync(htmlDocument);
        }

        /// <summary>
        /// Exports the css part of an <see cref="ExcelTable"/> to a html string
        /// </summary>
        /// <returns>A html table</returns>
        public Task<string> GetCssStringAsync()
        {
            var exporter = HtmlExporterFactory.CreateCssExporterAsync(_settings, _ranges);
            return exporter.GetCssStringAsync();
        }

        /// <summary>
        /// Exports the css part of an <see cref="ExcelTable"/> to a html string
        /// </summary>
        /// <returns>A html table</returns>
        public Task RenderCssAsync(Stream stream)
        {
            var exporter = HtmlExporterFactory.CreateCssExporterAsync(_settings, _ranges);
            return exporter.RenderCssAsync(stream);
        }
#endif
    }
}
