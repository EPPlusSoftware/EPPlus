/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  05/16/2020         EPPlus Software AB           ExcelTable Html Export
 *************************************************************************************************/
using OfficeOpenXml.Style.Table;
using OfficeOpenXml.Table;
using OfficeOpenXml.Utils;
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
    /// <summary>
    /// Exports a <see cref="ExcelTable"/> to Html
    /// </summary>
    public partial class TableExporter
    {
        /// <summary>
        /// Elements
        /// </summary>
        public Dictionary<string,string> AdditionalCssElements
        {
            get
            {
                return _genericCssElements;
            }
        }

        /// <summary>
        /// Exports an <see cref="ExcelTable"/> to a css string
        /// </summary>
        /// <returns>A cascading style sheet</returns>
        public string GetCssString()
        {
            return GetCssString(CssTableExportOptions.Default);
        }

        /// <summary>
        /// Exports an <see cref="ExcelTable"/> to a html string
        /// </summary>
        /// <param name="options"><see cref="HtmlTableExportOptions">Options</see> for the export</param>
        /// <returns>A html table</returns>
        public string GetCssString(CssTableExportOptions options)
        {
            using (var ms = new MemoryStream())
            {
                RenderCss(ms, options);
                ms.Position = 0;
                using (var sr = new StreamReader(ms))
                {
                    return sr.ReadToEnd();
                }
            }
        }

        public void RenderCss(Stream stream)
        {
            RenderCss(stream, CssTableExportOptions.Default);
        }

        public void RenderCss(Stream stream, Action<CssTableExportOptions> options)
        {
            var o = new CssTableExportOptions();
            options?.Invoke(o);
            RenderCss(stream, o);
        } 
        public void RenderCss(Stream stream, CssTableExportOptions options)
        {
            Require.Argument(options).IsNotNull("options");
            if (_table.TableStyle == TableStyles.None || options.IncludeTableStyles==false)
            {
                return; 
            }
            if (!stream.CanWrite)
            {
                throw new IOException("Parameter stream must be a writeable System.IO.Stream");
            }
            if (_datatypes.Count == 0) GetDataTypes(_table.Address);
            var writer = new EpplusTableCssWriter(stream, _table, options);
            writer.RenderAdditionalAndFontCss();
            if(options.IncludeTableStyles) writer.RenderTableCss(_datatypes);
            if(options.IncludeCellStyles) writer.RenderCellCss(_datatypes);
        }
    }
}
