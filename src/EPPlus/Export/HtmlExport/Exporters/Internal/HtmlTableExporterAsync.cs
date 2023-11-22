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
using OfficeOpenXml.Core.CellStore;
using OfficeOpenXml.Export.HtmlExport.Accessibility;
using OfficeOpenXml.Export.HtmlExport.Exporters.Internal;
using OfficeOpenXml.Export.HtmlExport.HtmlCollections;
using OfficeOpenXml.Export.HtmlExport.Parsers;
using OfficeOpenXml.Export.HtmlExport.Settings;
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
    internal class HtmlTableExporterAsync : HtmlTableExporterBase
    {
        internal HtmlTableExporterAsync(HtmlTableExportSettings settings, ExcelTable table) : base(settings, table, table.Range)
        {
        }

        /// <summary>
        /// Exports an <see cref="ExcelTable"/> to a html string
        /// </summary>
        /// <returns>A html table</returns>
        public async Task<string> GetHtmlStringAsync()
        {
            using (var ms = RecyclableMemory.GetStream())
            {
                await RenderHtmlAsync(ms);
                ms.Position = 0;
                using (var sr = new StreamReader(ms))
                {
                    return sr.ReadToEnd();
                }
            }
        }

        /// <summary>
        /// Exports the html part of an <see cref="ExcelTable"/> to a stream
        /// </summary>
        /// <returns>A html table</returns>
        public async Task RenderHtmlAsync(Stream stream)
        {
            ValidateStream(stream);
            var htmlTable = GenerateHtml();

            var writer = new HtmlWriter(stream, Settings.Encoding);
            await writer.RenderHTMLElementAsync(htmlTable, Settings.Minify);
        }

        /// <summary>
        /// Renders the Html and the Css to a single page. 
        /// </summary>
        /// <param name="htmlDocument">The html string where to insert the html and the css. The Html will be inserted in string parameter {0} and the Css will be inserted in parameter {1}.</param>
        /// <returns>The html document</returns>
        public async Task<string> GetSinglePageAsync(string htmlDocument = "<!DOCTYPE html>\r\n<html>\r\n<head>\r\n<style type=\"text/css\">\r\n{1}</style></head>\r\n<body>\r\n{0}</body>\r\n</html>")
        {
            if (Settings.Minify) htmlDocument = htmlDocument.Replace("\r\n", "");
            var html = await GetHtmlStringAsync();
            var cssExporter = HtmlExporterFactory.CreateCssExporterTableAsync(_tableExportSettings, _table, _exporterContext);
            var css = await cssExporter.GetCssStringAsync();
            return string.Format(htmlDocument, html, css);
        }
    }
}
#endif
