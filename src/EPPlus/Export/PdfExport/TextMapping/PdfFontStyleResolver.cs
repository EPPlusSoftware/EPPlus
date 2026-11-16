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
using OfficeOpenXml.Export.PdfExport.Data;
using OfficeOpenXml.Style.Dxf;
using OfficeOpenXml.Style.Table;
using OfficeOpenXml.Table;
using System.Collections.Generic;
using System.Linq;

namespace OfficeOpenXml.Export.PdfExport.TextMapping
{
    internal static class PdfFontStyleResolver
    {
        internal static PdfCellStyle GetFontStyle(ExcelRangeBase cell, PdfCellStyle cellStyle, Dictionary<ExcelTable, ExcelTableNamedStyle> tableStyleCache)
        {
            var cf = cell.ConditionalFormatting.GetConditionalFormattings();
            if (cf != null && cf.Count > 0)
            {
                // Sort ascending
                var ordered = cf.OrderBy(r => r.Priority);
                foreach (var rule in ordered)
                {
                    if (!rule.ShouldApplyToCell(cell))
                    {
                        if (rule.StopIfTrue) break;
                        continue;
                    }
                    if (rule.Style?.Font != null && rule.Style.Font.HasValue)
                    {
                        cellStyle.dxfFont = rule.Style.Font;
                        return cellStyle;
                    }
                    if (rule.StopIfTrue) break;
                }
            }
            var tables = cell.Worksheet.Tables.GetIntersectingRanges(cell);
            if (tables.Count > 0)
            {
                var table = tables[0].Value;
                var range = table.Range;
                ExcelTableNamedStyle tableStyle = PdfTableStyleLookup.GetTableStyle(table, tableStyleCache);
                if (tableStyle != null)
                {
                    int tableRow = cell._fromRow - range._fromRow;
                    int tableCol = cell._fromCol - range._fromCol;

                    ExcelDxfFontBase font = null;
                    foreach (var el in PdfTableStyleLookup.GetApplicableTableElements(table, tableStyle, tableRow, tableCol, cell))
                    {
                        var f = el.Style.Font;                                   // ExcelDxfFontBase
                        if (f == null || !f.HasValue) continue;
                        if (font == null) font = (ExcelDxfFontBase)f.Clone();   // Clone() returns DxfStyleBase
                        else OverlayFont(font, f);
                    }

                    PdfTableStyleLookup.GetTableRegionOverride(cell, table, out var ovCol, out var ovTbl);
                    cellStyle.dxfFontOverride = (ovCol?.Font != null && ovCol.Font.HasValue) ? ovCol.Font
                                              : (ovTbl?.Font != null && ovTbl.Font.HasValue) ? ovTbl.Font
                                              : null;
                    cellStyle.dxfFont = font;
                }
            }
            return cellStyle;
        }

        private static void OverlayFont(ExcelDxfFontBase dst, ExcelDxfFontBase src)
        {
            if (src.Bold != null) dst.Bold = src.Bold;
            if (src.Italic != null) dst.Italic = src.Italic;
            if (src.Strike != null) dst.Strike = src.Strike;
            if (src.Underline != null) dst.Underline = src.Underline;
            if (src.Color != null && src.Color.HasValue) dst.Color = src.Color;
        }
    }
}
