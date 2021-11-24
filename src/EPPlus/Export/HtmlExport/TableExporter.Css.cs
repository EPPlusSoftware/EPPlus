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
        /// Exports an <see cref="ExcelTable"/> to a html string
        /// </summary>
        /// <returns>A html table</returns>
        public string GetCssString()
        {
            using (var ms = new MemoryStream())
            {
                RenderCss(ms);
                ms.Position = 0;
                using (var sr = new StreamReader(ms))
                {
                    return sr.ReadToEnd();
                }
            }
        }
        public void RenderCss(Stream stream)
        {
            if ((_table.TableStyle == TableStyles.None || Settings.Css.IncludeTableStyles==false) && Settings.Css.IncludeCellStyles==false)
            {
                return; 
            }
            if (!stream.CanWrite)
            {
                throw new IOException("Parameter stream must be a writeable System.IO.Stream");
            }

            if (_datatypes.Count == 0) GetDataTypes(_table.Address);
            var sw = new StreamWriter(stream);
            if (Settings.Css.IncludeTableStyles) RenderTableCss(sw);
            if (Settings.Css.IncludeCellStyles) RenderCellCss(sw);
        }

        private void RenderCellCss(StreamWriter sw)
        {            
            var styleWriter = new EpplusCssWriter(sw, _table.Range, Settings);

            var r = _table.Range;
            var styles = r.Worksheet.Workbook.Styles;
            var ce = new CellStoreEnumerator<ExcelValue>(r.Worksheet._values, r._fromRow, r._fromCol, r._toRow, r._toCol);
            while (ce.Next())
            {
                if (ce.Value._styleId > 0 && ce.Value._styleId < styles.CellXfs.Count)
                {
                    styleWriter.AddToCss(styles, ce.Value._styleId);
                }
            }
            styleWriter.FlushStream();
        }

        internal void RenderTableCss(StreamWriter sw)
        {
            var styleWriter = new EpplusTableCssWriter(sw, _table, Settings);
            if (Settings.Minify == false) styleWriter.WriteLine();
            styleWriter.RenderAdditionalAndFontCss();
            ExcelTableNamedStyle tblStyle;
            if (_table.TableStyle == TableStyles.Custom)
            {
                tblStyle = _table.WorkSheet.Workbook.Styles.TableStyles[_table.StyleName].As.TableStyle;
            }
            else
            {
                var tmpNode = _table.WorkSheet.Workbook.StylesXml.CreateElement("c:tableStyle");
                tblStyle = new ExcelTableNamedStyle(_table.WorkSheet.Workbook.Styles.NameSpaceManager, tmpNode, _table.WorkSheet.Workbook.Styles);
                tblStyle.SetFromTemplate(_table.TableStyle);
            }

            var tableClass = $"{TableClass}.{TableExporter.TableStyleClassPrefix}{tblStyle.Name.ToLower()}";
            styleWriter.AddHyperlinkCss($"{tableClass}", tblStyle.WholeTable);
            styleWriter.AddAlignmentToCss($"{tableClass}", _datatypes);

            styleWriter.AddToCss($"{tableClass}", tblStyle.WholeTable, "");
            styleWriter.AddToCssBorderVH($"{tableClass}", tblStyle.WholeTable, "");

            //Header
            styleWriter.AddToCss($"{tableClass}", tblStyle.HeaderRow, " thead tr th");
            styleWriter.AddToCssBorderVH($"{tableClass}", tblStyle.HeaderRow, "");

            styleWriter.AddToCss($"{tableClass}", tblStyle.LastTotalCell, $" thead tr th:last-child)");
            styleWriter.AddToCss($"{tableClass}", tblStyle.FirstHeaderCell, " thead tr th:first-child");

            //Total
            styleWriter.AddToCss($"{tableClass}", tblStyle.TotalRow, " tfoot tr td");
            styleWriter.AddToCssBorderVH($"{tableClass}", tblStyle.TotalRow, "");
            styleWriter.AddToCss($"{tableClass}", tblStyle.LastTotalCell, $" tfoot tr td:last-child)");
            styleWriter.AddToCss($"{tableClass}", tblStyle.FirstTotalCell, " tfoot tr td:first-child");

            //Columns stripes
            tableClass = $"{TableExporter.TableStyleClassPrefix}{tblStyle.Name.ToLower()}-column-stripes";
            styleWriter.AddToCss($"{tableClass}", tblStyle.FirstColumnStripe, $" tbody tr td:nth-child(odd)");
            styleWriter.AddToCss($"{tableClass}", tblStyle.SecondColumnStripe, $" tbody tr td:nth-child(even)");

            //Row stripes
            tableClass = $"{TableExporter.TableStyleClassPrefix}{tblStyle.Name.ToLower()}-row-stripes";
            styleWriter.AddToCss($"{tableClass}", tblStyle.FirstRowStripe, " tbody tr:nth-child(odd)");
            styleWriter.AddToCss($"{tableClass}", tblStyle.SecondRowStripe, " tbody tr:nth-child(even)");

            //Last column
            tableClass = $"{TableExporter.TableStyleClassPrefix}{tblStyle.Name.ToLower()}-last-column";
            styleWriter.AddToCss($"{tableClass}", tblStyle.LastColumn, $" tbody tr td:last-child");

            //First column
            tableClass = $"{TableExporter.TableStyleClassPrefix}{tblStyle.Name.ToLower()}-first-column";
            styleWriter.AddToCss($"{tableClass}", tblStyle.FirstColumn, " tbody tr td:first-child");

            styleWriter.FlushStream();
        }
    }
}
