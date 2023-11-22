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
using OfficeOpenXml.Export.HtmlExport.Writers;
using OfficeOpenXml.Table;
using OfficeOpenXml.Utils;
using System.IO;

namespace OfficeOpenXml.Export.HtmlExport.Exporters.Internal
{
    internal class CssRangeExporterSync : CssExporterBase
    {
        public CssRangeExporterSync(HtmlRangeExportSettings settings, EPPlusReadOnlyList<ExcelRangeBase> ranges)
            : base(settings, ranges)
        {
            _settings = settings;
        }

        public CssRangeExporterSync(HtmlRangeExportSettings settings, ExcelRangeBase range)
            : base(settings, range)
        {
            _settings = settings;
        }

        HtmlRangeExportSettings _settings;

        /// <summary>
        /// Exports an <see cref="ExcelTable"/> to a html string
        /// </summary>
        /// <returns>A html table</returns>
        public string GetCssString()
        {
            using (var ms = RecyclableMemory.GetStream())
            {
                RenderCss(ms);
                ms.Position = 0;
                using (var sr = new StreamReader(ms))
                {
                    return sr.ReadToEnd();
                }
            }
        }
        /// <summary>
        /// Exports the css part of the html export.
        /// </summary>
        /// <param name="stream">The stream to write the css to.</param>
        /// <exception cref="IOException"></exception>
        public void RenderCss(Stream stream)
        {
            var trueWriter = new CssWriter(stream);
            var cssRules = CreateRuleCollection(_settings);

            trueWriter.WriteAndClearFlush(cssRules, Settings.Minify);
        }
    }
}
