using OfficeOpenXml.Core.CellStore;
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

namespace OfficeOpenXml.Export.HtmlExport.Exporters
{
    internal class CssTableExporterAsync : CssRangeExporterBase
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
            var ranges = new List<ExcelRangeBase>() { _table.Range };
            var cellCssWriter = new EpplusCssWriter(sw, ranges, _tableSettings, _tableSettings.Css, _tableSettings.Css.Exclude.CellStyle, _styleCache);
            await cellCssWriter.RenderAdditionalAndFontCssAsync(TableClass);
            if (_tableSettings.Css.IncludeTableStyles) await RenderTableCssAsync(sw, _table, _tableSettings, _styleCache, _dataTypes);
            if (_tableSettings.Css.IncludeCellStyles) await RenderCellCssAsync(sw);

            if (Settings.Pictures.Include == ePictureInclude.Include)
            {
                LoadRangeImages(ranges);
                foreach (var p in _rangePictures)
                {
                    await cellCssWriter.AddPictureToCssAsync(p);
                }
            }
            await cellCssWriter.FlushStreamAsync();
        }

        private async Task RenderCellCssAsync(StreamWriter sw)
        {
            var ranges = new List<ExcelRangeBase>() { _table.Range };
            var styleWriter = new EpplusCssWriter(sw, ranges, _tableSettings, _tableSettings.Css, _tableSettings.Css.Exclude.CellStyle, _styleCache);

            var r = _table.Range;
            var styles = r.Worksheet.Workbook.Styles;
            var ce = new CellStoreEnumerator<ExcelValue>(r.Worksheet._values, r._fromRow, r._fromCol, r._toRow, r._toCol);
            while (ce.Next())
            {
                if (ce.Value._styleId > 0 && ce.Value._styleId < styles.CellXfs.Count)
                {
                    await styleWriter.AddToCssAsync(styles, ce.Value._styleId, Settings.StyleClassPrefix, Settings.CellStyleClassName);
                }
            }
            await styleWriter.FlushStreamAsync();
        }

        internal async Task RenderTableCssAsync(StreamWriter sw, ExcelTable table, HtmlTableExportSettings settings, Dictionary<string, int> styleCache, List<string> datatypes)
        {
            var styleWriter = new EpplusTableCssWriter(sw, table, settings, styleCache);
            if (settings.Minify == false) await styleWriter.WriteLineAsync();
            ExcelTableNamedStyle tblStyle;
            if (table.TableStyle == TableStyles.Custom)
            {
                tblStyle = table.WorkSheet.Workbook.Styles.TableStyles[table.StyleName].As.TableStyle;
            }
            else
            {
                var tmpNode = table.WorkSheet.Workbook.StylesXml.CreateElement("c:tableStyle");
                tblStyle = new ExcelTableNamedStyle(table.WorkSheet.Workbook.Styles.NameSpaceManager, tmpNode, table.WorkSheet.Workbook.Styles);
                tblStyle.SetFromTemplate(table.TableStyle);
            }

            var tableClass = $"{TableClass}.{HtmlExportTableUtil.TableStyleClassPrefix}{GetClassName(tblStyle.Name, "EmptyClassName").ToLower()}";
            await styleWriter.AddHyperlinkCssAsync($"{tableClass}", tblStyle.WholeTable);
            await styleWriter.AddAlignmentToCssAsync($"{tableClass}", datatypes);

            await styleWriter.AddToCssAsync($"{tableClass}", tblStyle.WholeTable, "");
            await styleWriter.AddToCssBorderVHAsync($"{tableClass}", tblStyle.WholeTable, "");

            //Header
            await styleWriter.AddToCssAsync($"{tableClass}", tblStyle.HeaderRow, " thead");
            await styleWriter.AddToCssBorderVHAsync($"{tableClass}", tblStyle.HeaderRow, "");

            await styleWriter.AddToCssAsync($"{tableClass}", tblStyle.LastTotalCell, $" thead tr th:last-child)");
            await styleWriter.AddToCssAsync($"{tableClass}", tblStyle.FirstHeaderCell, " thead tr th:first-child");

            //Total
            await styleWriter.AddToCssAsync($"{tableClass}", tblStyle.TotalRow, " tfoot");
            await styleWriter.AddToCssBorderVHAsync($"{tableClass}", tblStyle.TotalRow, "");
            await styleWriter.AddToCssAsync($"{tableClass}", tblStyle.LastTotalCell, $" tfoot tr td:last-child)");
            await styleWriter.AddToCssAsync($"{tableClass}", tblStyle.FirstTotalCell, " tfoot tr td:first-child");

            //Columns stripes
            var tableClassCS = $"{tableClass}-column-stripes";
            await styleWriter.AddToCssAsync($"{tableClassCS}", tblStyle.FirstColumnStripe, $" tbody tr td:nth-child(odd)");
            await styleWriter.AddToCssAsync($"{tableClassCS}", tblStyle.SecondColumnStripe, $" tbody tr td:nth-child(even)");

            //Row stripes
            var tableClassRS = $"{tableClass}-row-stripes";
            await styleWriter.AddToCssAsync($"{tableClassRS}", tblStyle.FirstRowStripe, " tbody tr:nth-child(odd)");
            await styleWriter.AddToCssAsync($"{tableClassRS}", tblStyle.SecondRowStripe, " tbody tr:nth-child(even)");

            //Last column
            var tableClassLC = $"{tableClass}-last-column";
            await styleWriter.AddToCssAsync($"{tableClassLC}", tblStyle.LastColumn, $" tbody tr td:last-child");

            //First column
            var tableClassFC = $"{tableClass}-first-column";
            await styleWriter.AddToCssAsync($"{tableClassFC}", tblStyle.FirstColumn, " tbody tr td:first-child");


            await styleWriter.FlushStreamAsync();
        }
    }
}
#endif
