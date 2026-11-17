/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  27/11/2025         EPPlus Software AB           EPPlus 9
 *************************************************************************************************/
using OfficeOpenXml.Style;
using OfficeOpenXml.Style.Dxf;
using OfficeOpenXml.Style.Table;
using OfficeOpenXml.Table;
using System.Collections.Generic;

namespace OfficeOpenXml.Export.PdfExport.TextMapping
{
    internal static class PdfTableStyleLookup
    {
        internal static ExcelTableNamedStyle GetTableStyle(ExcelTable table, Dictionary<ExcelTable, ExcelTableNamedStyle> cache)
        {
            if (cache.TryGetValue(table, out var cached))
                return cached;
            ExcelTableNamedStyle tableStyle;
            if (table.TableStyle == TableStyles.Custom)
            {
                tableStyle = table.WorkSheet.Workbook.Styles.TableStyles[table.StyleName].As.TableStyle;
            }
            else
            {
                var tmpNode = table.WorkSheet.Workbook.StylesXml.CreateElement("c:tableStyle");
                tableStyle = new ExcelTableNamedStyle(
                    table.WorkSheet.Workbook.Styles.NameSpaceManager, tmpNode, table.WorkSheet.Workbook.Styles);
                tableStyle.SetFromTemplate((TableStyles)table.TableStyle);
            }
            cache[table] = tableStyle;
            return tableStyle;
        }

        internal static IEnumerable<ExcelTableStyleElement> GetApplicableTableElements(ExcelTable table, ExcelTableNamedStyle ts, int tableRow, int tableCol, ExcelRangeBase cell)
        {
            var range = table.Range;
            bool firstCol = table.ShowFirstColumn && tableCol == 0;
            bool lastCol = table.ShowLastColumn && cell._fromCol == range._toCol;
            bool header = table.ShowHeader && tableRow == 0;
            bool total = table.ShowTotal && cell._fromRow == range._toRow;
            yield return ts.WholeTable;                                                 // 1
            // column-spanning: header + data + total
            if (table.ShowColumnStripes)
                yield return (tableCol & 1) == 0 ? ts.FirstColumnStripe                 // 2
                                                 : ts.SecondColumnStripe;               // 3
                                                                                        // row stripes: data rows only
            if (!header && !total && table.ShowRowStripes)
                yield return (tableRow & 1) != 0 ? ts.FirstRowStripe                    // 4
                                                 : ts.SecondRowStripe;                  // 5

            if (lastCol) yield return ts.LastColumn;                                    // 6
            if (firstCol) yield return ts.FirstColumn;                                  // 7
            if (total) yield return ts.TotalRow;                                        // 8
            if (header) yield return ts.HeaderRow;                                      // 9
            if (header && firstCol) yield return ts.FirstHeaderCell;                    // 10
            if (header && lastCol) yield return ts.LastHeaderCell;                      // 11
            if (total && firstCol) yield return ts.FirstTotalCell;                      // 12
            if (total && lastCol) yield return ts.LastTotalCell;                        // 13
        }

        internal static void GetTableRegionOverride(ExcelRangeBase cell, ExcelTable table, out ExcelDxfStyle colStyle, out ExcelDxfStyle tblStyle)
        {
            colStyle = null;
            tblStyle = null;
            var range = table.Range;
            int tableRow = cell._fromRow - range._fromRow;
            int tableCol = cell._fromCol - range._fromCol;
            ExcelTableColumn column = (tableCol >= 0 && tableCol < table.Columns.Count) ? table.Columns[tableCol] : null;

            string attr;
            ExcelDxfStyle tblCandidate, colCandidate;
            if (table.ShowHeader && tableRow == 0)
            { attr = "headerRowDxfId"; tblCandidate = table.HeaderRowStyle; colCandidate = column?.HeaderRowStyle; }
            else if (table.ShowTotal && cell._fromRow == range._toRow)
            { attr = "totalsRowDxfId"; tblCandidate = table.TotalsRowStyle; colCandidate = column?.TotalsRowStyle; }
            else
            { attr = "dataDxfId"; tblCandidate = table.DataStyle; colCandidate = column?.DataStyle; }

            var tableNode = table.TableXml?.DocumentElement;
            if (tableNode == null) return;

            if (tableNode.Attributes?[attr] != null) tblStyle = tblCandidate;

            if (column != null)
            {
                var colNode = GetTableColumnNode(tableNode, tableCol);
                if (colNode?.Attributes?[attr] != null) colStyle = colCandidate;
            }
        }

        internal static System.Xml.XmlNode GetTableColumnNode(System.Xml.XmlNode tableNode, int index)
        {
            foreach (System.Xml.XmlNode child in tableNode.ChildNodes)
            {
                if (child.LocalName != "tableColumns") continue;
                int i = 0;
                foreach (System.Xml.XmlNode col in child.ChildNodes)
                {
                    if (col.LocalName != "tableColumn") continue;
                    if (i == index) return col;
                    i++;
                }
            }
            return null;
        }
    }
}
