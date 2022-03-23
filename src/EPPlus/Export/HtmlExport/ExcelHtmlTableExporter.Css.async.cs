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
using OfficeOpenXml.Core.CellStore;
using OfficeOpenXml.Style.Table;
using OfficeOpenXml.Table;
using OfficeOpenXml.Utils;
using System.Collections.Generic;
using System.IO;
#if !NET35 && !NET40
using System.Threading.Tasks;
#endif

namespace OfficeOpenXml.Export.HtmlExport
{
#if !NET35 && !NET40
    /// <summary>
    /// Exports a <see cref="ExcelTable"/> to Html
    /// </summary>
    public partial class ExcelHtmlTableExporter
    {        
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
            if ((_table.TableStyle == TableStyles.None || Settings.Css.IncludeTableStyles==false) && Settings.Css.IncludeCellStyles==false)
            {
                return; 
            }
            if (!stream.CanWrite)
            {
                throw new IOException("Parameter stream must be a writeable System.IO.Stream");
            }

            if (_dataTypes.Count == 0) GetDataTypes(_table.Address);
            var sw = new StreamWriter(stream);
            var ranges = new List<ExcelRangeBase>() { _table.Range };
            var cellCssWriter = new EpplusCssWriter(sw, ranges, Settings, Settings.Css, Settings.Css.Exclude.CellStyle, _styleCache);
            await cellCssWriter.RenderAdditionalAndFontCssAsync(TableClass);
            if (Settings.Css.IncludeTableStyles) await RenderTableCssAsync(sw, _table, Settings, _styleCache, _dataTypes);
            if (Settings.Css.IncludeCellStyles) await RenderCellCssAsync(sw);

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
            var styleWriter = new EpplusCssWriter(sw, ranges, Settings, Settings.Css, Settings.Css.Exclude.CellStyle, _styleCache);

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

        internal static async Task RenderTableCssAsync(StreamWriter sw, ExcelTable table, HtmlTableExportSettings settings, Dictionary<string, int> styleCache, List<string> datatypes)
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

            var tableClass = $"{TableClass}.{TableStyleClassPrefix}{GetClassName(tblStyle.Name, "EmptyClassName").ToLower()}";
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
#endif
}
