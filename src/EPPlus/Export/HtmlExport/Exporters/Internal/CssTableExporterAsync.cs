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
using OfficeOpenXml.Export.HtmlExport.Exporters.Internal;
using OfficeOpenXml.Export.HtmlExport.Settings;
using OfficeOpenXml.Table;
using OfficeOpenXml.Utils;
using System.IO;
#if !NET35 && !NET40
using System.Threading.Tasks;

namespace OfficeOpenXml.Export.HtmlExport.Exporters
{
    internal class CssTableExporterAsync : CssExporterBase
    {
        public CssTableExporterAsync(HtmlTableExportSettings settings, ExcelTable table) : base(settings, table.Range)
        {
            _table = table;
            _tableSettings = settings;
        }

        private readonly ExcelTable _table;
        private readonly HtmlTableExportSettings _tableSettings;

        /// <summary>
        /// Exports the css part of an <see cref="ExcelTable"/> to a html string
        /// </summary>
        /// <returns>A html table</returns>
        public async Task<string> GetCssStringAsync()
        {
            using (var ms = RecyclableMemory.GetStream())
            {
                await RenderCssAsync(ms);
                ms.Position = 0;
                using (var sr = new StreamReader(ms))
                {
                    return await sr.ReadToEndAsync();
                }
            }
        }
        /// <summary>
        /// Exports the css part of an <see cref="ExcelTable"/> to a html string
        /// </summary>
        /// <returns>A html table</returns>
        public async Task RenderCssAsync(Stream stream)
        {
            var cssWriter = GetTableCssWriter(stream, _table, _tableSettings);
            if (cssWriter == null) { return; }

            var cssRules = CreateRuleCollection(_tableSettings);
            await cssWriter.WriteAndClearFlushAsync(cssRules, Settings.Minify);
        }
    }
}
#endif
