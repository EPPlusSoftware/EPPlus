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
using OfficeOpenXml.Utils;
using OfficeOpenXml.Table;
using System.IO;
using OfficeOpenXml.Export.HtmlExport.Settings;

namespace OfficeOpenXml.Export.HtmlExport.Exporters.Internal
{
    internal class HtmlTableExporterSync : HtmlTableExporterBase
    {
        internal HtmlTableExporterSync(HtmlTableExportSettings settings, ExcelTable table)
            : base(settings, table, table.Range)
        {}

        /// <summary>
        /// Exports an <see cref="ExcelTable"/> to a html string
        /// </summary>
        /// <returns>A html table</returns>
        public string GetHtmlString()
        {
            using (var ms = RecyclableMemory.GetStream())
            {
                RenderHtml(ms);
                ms.Position = 0;
                using (var sr = new StreamReader(ms))
                {
                    return sr.ReadToEnd();
                }
            }
        }
        /// <summary>
        /// Exports the html part of an <see cref="ExcelTable"/> to a html string.
        /// </summary>
        /// <param name="stream">The stream to write to.</param>
        /// <exception cref="IOException"></exception>
        public void RenderHtml(Stream stream)
        {
            ValidateStream(stream);
            var htmlTable = GenerateHtml();

            var writer = new HtmlWriter(stream, Settings.Encoding);
            writer.RenderHTMLElement(htmlTable, Settings.Minify);
        }

        /// <summary>
        /// Renders both the Css and the Html to a single page. 
        /// </summary>
        /// <param name="htmlDocument">The html string where to insert the html and the css. The Html will be inserted in string parameter {0} and the Css will be inserted in parameter {1}.</param>
        /// <returns>The html document</returns>
        public string GetSinglePage(string htmlDocument = "<!DOCTYPE html>\r\n<html>\r\n<head>\r\n<style type=\"text/css\">\r\n{1}</style></head>\r\n<body>\r\n{0}</body>\r\n</html>")
        {
            if (Settings.Minify) htmlDocument = htmlDocument.Replace("\r\n", "");
            var html = GetHtmlString();
            var cssExporter = HtmlExporterFactory.CreateCssExporterTableSync(_tableExportSettings, _table, _exporterContext);
            var css = cssExporter.GetCssString();
            return string.Format(htmlDocument, html, css);

        }
    }
}
