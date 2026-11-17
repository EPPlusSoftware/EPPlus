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
using EPPlus.Export.Pdf.Enums;
using OfficeOpenXml.ConditionalFormatting;
using OfficeOpenXml.ConditionalFormatting.Contracts;
using OfficeOpenXml.Export.PdfExport.Data;
using OfficeOpenXml.Style.Dxf;
using OfficeOpenXml.Style.Table;
using OfficeOpenXml.Table;
using System.Collections.Generic;
using System.Linq;

namespace OfficeOpenXml.Export.PdfExport.TextMapping
{
    internal static class PdfFillStyleResolver
    {
        internal static void GetFillStyles(ExcelRangeBase cell, PdfCellStyle cellStyle, Dictionary<ExcelTable, ExcelTableNamedStyle> tableStyleCache)
        {
            cellStyle.xfFill = cell.Style.Fill;

            // Conditional formatting is evaluated regardless of the cell's own fill, because a
            // color scale paints over the base cell background (Excel behaviour).
            var cf = cell.ConditionalFormatting.GetConditionalFormattings();
            if (cf != null && cf.Count > 0)
            {
                var ordered = cf.OrderBy(r => r.Priority);
                foreach (var rule in ordered)
                {
                    if (!rule.ShouldApplyToCell(cell))
                    {
                        if (rule.StopIfTrue) break;
                        continue;
                    }
                    //Data Bar
                    if (rule is ExcelConditionalFormattingDataBar)
                    {
                        if (rule.StopIfTrue) break;
                        continue;
                    }
                    //Icon set
                    if (rule is IExcelConditionalFormattingIconSet)
                    {
                        if (rule.StopIfTrue) break;
                        continue;
                    }
                    // Color scale (two/three)
                    if (rule is ExcelConditionalFormattingTwoColorScale colorScale)
                    {
                        var c = colorScale.GetColorScaleColor(new ExcelAddress(cell.Address));
                        if (c != null)
                        {
                            cellStyle.ConditionalFill = c;
                            return;
                        }
                        if (rule.StopIfTrue) break;
                        continue;
                    }
                    // dxf fill
                    if (rule.Style?.Fill != null && rule.Style.Fill.HasValue)
                    {
                        cellStyle.dxfFill = rule.Style.Fill;
                        return;
                    }

                    if (rule.StopIfTrue) break;
                }
            }

            // Table styling only applies to cells without an explicit fill of their own.
            if (cell.Style.Fill.IsEmpty())
            {
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

                        ExcelDxfFill fill = null;
                        foreach (var el in PdfTableStyleLookup.GetApplicableTableElements(table, tableStyle, tableRow, tableCol, cell))
                        {
                            var f = el.Style.Fill;
                            if (FillIsPaintable(f)) fill = f;
                        }

                        PdfTableStyleLookup.GetTableRegionOverride(cell, table, out var ovCol, out var ovTbl);
                        var ovFill = (ovCol?.Fill != null && ovCol.Fill.HasValue) ? ovCol.Fill
                                   : (ovTbl?.Fill != null && ovTbl.Fill.HasValue) ? ovTbl.Fill
                                   : null;
                        if (ovFill != null && FillIsPaintable(ovFill)) fill = ovFill;

                        cellStyle.dxfFill = fill;
                    }
                }
            }
        }

        private static bool FillIsPaintable(ExcelDxfFill fill)
        {
            if (fill == null || !fill.HasValue)
                return false;
            var pattern = fill.PatternType != null ? (ExcelFillStyle)fill.PatternType : ExcelFillStyle.Solid;
            if (pattern == ExcelFillStyle.None)
                return fill.Gradient != null;
            if (pattern == ExcelFillStyle.Solid)
                return !string.IsNullOrEmpty(fill.BackgroundColor?.LookupColor());
            return true;
        }
    }
}
