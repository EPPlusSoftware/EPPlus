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

namespace OfficeOpenXml.Export.HtmlExport.Exporters
{
    internal class ExcelHtmlTableExporter : IExcelHtmlTableExporter
    {
        public ExcelHtmlTableExporter(ExcelTable table)
        {
            _table = table;
            _settings = new HtmlTableExportSettings();
        }

        private readonly ExcelTable _table;
        private readonly HtmlTableExportSettings _settings;
        private readonly Dictionary<string, int> _styleCache = new Dictionary<string, int>();

        public HtmlTableExportSettings Settings
        { get { return _settings; } }

        /// <summary>
        /// Exports an <see cref="ExcelTable"/> to a html string
        /// </summary>
        /// <returns>A html table</returns>
        public string GetHtmlString()
        {
            var exporter = HtmlExporterFactory.CreateHtmlTableExporterSync(_settings, _table, _styleCache);
            return exporter.GetHtmlString();
        }

        /// <summary>
        /// Exports an <see cref="ExcelTable"/> to a html string
        /// </summary>
        /// <param name="stream">The stream to write to</param>
        /// <returns>A html table</returns>
        public void RenderHtml(Stream stream)
        {
            var exporter = HtmlExporterFactory.CreateHtmlTableExporterSync(_settings, _table, _styleCache);
            exporter.RenderHtml(stream);
        }

        /// <summary>
        /// Renders both the Html and the Css to a single page. 
        /// </summary>
        /// <param name="htmlDocument">The html string where to insert the html and the css. The Html will be inserted in string parameter {0} and the Css will be inserted in parameter {1}.</param>
        /// <returns>The html document</returns>
        public string GetSinglePage(string htmlDocument = "<!DOCTYPE html>\r\n<html>\r\n<head>\r\n<style type=\"text/css\">\r\n{1}</style></head>\r\n<body>\r\n{0}</body>\r\n</html>")
        {
            var exporter = HtmlExporterFactory.CreateHtmlTableExporterSync(_settings, _table, _styleCache);
            return exporter.GetSinglePage(htmlDocument);
        }

        /// <summary>
        /// Exports an <see cref="ExcelTable"/> to a html string
        /// </summary>
        /// <returns>Cascading style sheet for the exported range</returns>
        public string GetCssString()
        {
            var exporter = HtmlExporterFactory.CreateCssExporterTableSync(_settings, _table, _styleCache);
            return exporter.GetCssString();
        }

        /// <summary>
        /// Exports the css part of the html export.
        /// </summary>
        /// <param name="stream">The stream to write the css to.</param>
        /// <exception cref="IOException"></exception>
        public void RenderCss(Stream stream)
        {
            var exporter = HtmlExporterFactory.CreateCssExporterTableSync(_settings, _table, _styleCache);
            exporter.RenderCss(stream);
        }

#if !NET35 && !NET40
        /// <summary>
        /// Exports an <see cref="ExcelTable"/> to a html string
        /// </summary>
        /// <returns>A html table</returns>
        public Task<string> GetHtmlStringAsync()
        {
            var exporter = HtmlExporterFactory.CreateHtmlTableExporterAsync(_settings, _table, _styleCache);
            return exporter.GetHtmlStringAsync();
        }

        /// <summary>
        /// Exports an <see cref="ExcelTable"/> to a html string
        /// </summary>
        /// <param name="stream">The stream to write to</param>
        /// <returns>A html table</returns>
        public Task RenderHtmlAsync(Stream stream)
        {
            var exporter = HtmlExporterFactory.CreateHtmlTableExporterAsync(_settings, _table, _styleCache);
            return exporter.RenderHtmlAsync(stream);
        }

        /// <summary>
        /// Renders the first range of the Html and the Css to a single page. 
        /// </summary>
        /// <param name="htmlDocument">The html string where to insert the html and the css. The Html will be inserted in string parameter {0} and the Css will be inserted in parameter {1}.</param>
        /// <returns>The html document</returns>
        public Task<string> GetSinglePageAsync(string htmlDocument = "<!DOCTYPE html>\r\n<html>\r\n<head>\r\n<style type=\"text/css\">\r\n{1}</style></head>\r\n<body>\r\n{0}</body>\r\n</html>")
        {
            var exporter = HtmlExporterFactory.CreateHtmlTableExporterAsync(_settings, _table, _styleCache);
            return exporter.GetSinglePageAsync(htmlDocument);
        }

        /// <summary>
        /// Exports the css part of an <see cref="ExcelTable"/> to a html string
        /// </summary>
        /// <returns>A html table</returns>
        public Task<string> GetCssStringAsync()
        {
            var exporter = HtmlExporterFactory.CreateCssExporterTableAsync(_settings, _table, _styleCache);
            return exporter.GetCssStringAsync();
        }

        /// <summary>
        /// Exports the css part of an <see cref="ExcelTable"/> to a html string
        /// </summary>
        /// <returns>A html table</returns>
        public Task RenderCssAsync(Stream stream)
        {
            var exporter = HtmlExporterFactory.CreateCssExporterTableAsync(_settings, _table, _styleCache);
            return exporter.RenderCssAsync(stream);
        }
#endif
    }
}
