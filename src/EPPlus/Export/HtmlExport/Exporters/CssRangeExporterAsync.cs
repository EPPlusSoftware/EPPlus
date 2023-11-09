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
using OfficeOpenXml.Export.HtmlExport.Writers.Css;
using OfficeOpenXml.Export.HtmlExport.Writers;
using OfficeOpenXml.Table;
using OfficeOpenXml.Utils;
using System.IO;
using System.Linq;
using System.Runtime;
#if !NET35 && !NET40
using System.Threading.Tasks;

namespace OfficeOpenXml.Export.HtmlExport.Exporters
{
    internal class CssRangeExporterAsync : CssExporterBase
    {
        public CssRangeExporterAsync(HtmlRangeExportSettings settings, EPPlusReadOnlyList<ExcelRangeBase> ranges)
         : base(settings, ranges)
        { _settings = settings;  }

        public CssRangeExporterAsync(HtmlRangeExportSettings settings, ExcelRangeBase range)
            : base(settings, range)
        { _settings = settings; }

        HtmlRangeExportSettings _settings;

        /// <summary>
        /// Exports an <see cref="ExcelTable"/> to a html string
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
        /// Exports an <see cref="ExcelTable"/> to html and writes it to a stream
        /// </summary>
        /// <param name="stream">The stream to write to</param>
        /// <returns></returns>
        public async Task RenderCssAsync(Stream stream)
        {
            if (!stream.CanWrite)
            {
                throw new IOException("Parameter stream must be a writable System.IO.Stream");
            }

            var sw = new StreamWriter(stream);
            await WriteCellAsync(sw);
        }

        private async Task WriteCellAsync(StreamWriter sw)
        {
            var trueWriter = new CssTrueWriter(sw);
            var cssRules = CreateRuleCollection(_settings);

            await trueWriter.WriteAndClearCollectionAsync(cssRules, Settings.Minify);
            await sw.FlushAsync();
        }
    }
}
#endif
