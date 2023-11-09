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
using OfficeOpenXml.Export.HtmlExport.Settings;
using OfficeOpenXml.Export.HtmlExport.Writers;
using OfficeOpenXml.Table;
using OfficeOpenXml.Utils;
using System.IO;

namespace OfficeOpenXml.Export.HtmlExport.Exporters
{
    internal class CssTableExporterSync : CssExporterBase
    {
        public CssTableExporterSync(HtmlTableExportSettings settings, ExcelTable table) : base(settings, table.Range)
        {
            _table = table;
            _tableSettings = settings;
        }

        private readonly ExcelTable _table;
        private readonly HtmlTableExportSettings _tableSettings;

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
        /// Exports the css part of an <see cref="ExcelTable"/> to a html string
        /// </summary>
        /// <returns>A html table</returns>
        public void RenderCss(Stream stream)
        {
            if ((_table.TableStyle == TableStyles.None || _tableSettings.Css.IncludeTableStyles == false) && _tableSettings.Css.IncludeCellStyles == false)
            {
                return;
            }
            if (!stream.CanWrite)
            {
                throw new IOException("Parameter stream must be a writeable System.IO.Stream");
            }

            if (_dataTypes.Count == 0) GetDataTypes(_table.Address, _table);

            var sw = new StreamWriter(stream);
            var trueWriter = new CssTrueWriter(sw);

            var cssRules = CreateRuleCollection(_tableSettings);
            trueWriter.WriteAndClearCollection(cssRules, Settings.Minify);
            sw.Flush();
        }
    }
}
