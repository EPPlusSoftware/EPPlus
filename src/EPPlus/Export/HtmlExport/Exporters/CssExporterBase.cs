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
using OfficeOpenXml.ConditionalFormatting;
using OfficeOpenXml.Core;
using OfficeOpenXml.Core.RangeQuadTree;
using OfficeOpenXml.Export.HtmlExport.Collectors;
using OfficeOpenXml.Export.HtmlExport.Settings;
using OfficeOpenXml.Style.Table;
using OfficeOpenXml.Table;
using OfficeOpenXml.Utils;
using System.Collections.Generic;

namespace OfficeOpenXml.Export.HtmlExport.Exporters
{
    internal abstract class CssExporterBase : AbstractHtmlExporter
    {
        public CssExporterBase(HtmlExportSettings settings, ExcelRangeBase range)
        {
            Settings = settings;
            Require.Argument(range).IsNotNull("range");
            _ranges = new EPPlusReadOnlyList<ExcelRangeBase>();

            if (range.Addresses == null)
            {
                AddRange(range);
            }
            else
            {
                foreach (var address in range.Addresses)
                {
                    AddRange(range.Worksheet.Cells[address.Address]);
                }
            }
        }

        public CssExporterBase(HtmlRangeExportSettings settings, EPPlusReadOnlyList<ExcelRangeBase> ranges)
        {
            Settings = settings;
            Require.Argument(ranges).IsNotNull("ranges");
            _ranges = ranges;
        }

        protected HtmlExportSettings Settings;
        protected EPPlusReadOnlyList<ExcelRangeBase> _ranges = new EPPlusReadOnlyList<ExcelRangeBase>();
        internal const string TableStyleClassPrefix = "ts-";

        private void AddRange(ExcelRangeBase range)
        {
            if (range.IsFullColumn && range.IsFullRow)
            {
                _ranges.Add(new ExcelRangeBase(range.Worksheet, range.Worksheet.Dimension.Address));
            }
            else
            {
                _ranges.Add(range);
            }
        }

        internal static CssTableRuleCollection RenderTableCss(ExcelTable table, HtmlTableExportSettings settings, List<string> datatypes)
        {
            var tableRules = new CssTableRuleCollection(table, settings);

            //if (settings.Minify == false) styleWriter.WriteLine();
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

            var tableClass = $"{TableClass}.{TableStyleClassPrefix}{HtmlExportTableUtil.GetClassName(tblStyle.Name, "EmptyTableStyle").ToLower()}";

            tableRules.AddHyperlink($"{tableClass}", tblStyle.WholeTable);
            tableRules.AddAlignment($"{tableClass}", datatypes);

            tableRules.AddToCollection($"{tableClass}", tblStyle.WholeTable, "");
            tableRules.AddToCollectionVH($"{tableClass}", tblStyle.WholeTable, "");

            //Header
            tableRules.AddToCollection($"{tableClass}", tblStyle.HeaderRow, " thead");
            tableRules.AddToCollectionVH($"{tableClass}", tblStyle.HeaderRow, "");

            tableRules.AddToCollection($"{tableClass}", tblStyle.LastTotalCell, $" thead tr th:last-child)");
            tableRules.AddToCollection($"{tableClass}", tblStyle.FirstHeaderCell, " thead tr th:first-child");

            //Total
            tableRules.AddToCollection($"{tableClass}", tblStyle.TotalRow, " tfoot");
            tableRules.AddToCollectionVH($"{tableClass}", tblStyle.TotalRow, "");
            tableRules.AddToCollection($"{tableClass}", tblStyle.LastTotalCell, $" tfoot tr td:last-child)");
            tableRules.AddToCollection($"{tableClass}", tblStyle.FirstTotalCell, " tfoot tr td:first-child");

            //Columns stripes
            var tableClassCS = $"{tableClass}-column-stripes";
            tableRules.AddToCollection($"{tableClassCS}", tblStyle.FirstColumnStripe, $" tbody tr td:nth-child(odd)");
            tableRules.AddToCollection($"{tableClassCS}", tblStyle.SecondColumnStripe, $" tbody tr td:nth-child(even)");

            //Row stripes
            var tableClassRS = $"{tableClass}-row-stripes";
            tableRules.AddToCollection($"{tableClassRS}", tblStyle.FirstRowStripe, " tbody tr:nth-child(odd)");
            tableRules.AddToCollection($"{tableClassRS}", tblStyle.SecondRowStripe, " tbody tr:nth-child(even)");

            //Last column
            var tableClassLC = $"{tableClass}-last-column";
            tableRules.AddToCollection($"{tableClassLC}", tblStyle.LastColumn, $" tbody tr td:last-child");

            //First column
            var tableClassFC = $"{tableClass}-first-column";
            tableRules.AddToCollection($"{tableClassFC}", tblStyle.FirstColumn, " tbody tr td:first-child");

            return tableRules;
        }
    }
}
