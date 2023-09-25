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
using OfficeOpenXml.Style.Table;
using OfficeOpenXml.Table;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
#if !NET35 && !NET40
using System.Threading.Tasks;
#endif

namespace OfficeOpenXml.Export.HtmlExport.Exporters
{
    internal static class HtmlExportTableUtil
    {

        internal const string TableStyleClassPrefix = "ts-";
        internal const string TableClass = "epplus-table";

        internal static string GetClassName(string className, string optionalName)
        {
            if (string.IsNullOrEmpty(optionalName)) return optionalName;

            className = className.Trim().Replace(" ", "-");
            var newClassName = "";
            for (int i = 0; i < className.Length; i++)
            {
                var c = className[i];
                if (i == 0)
                {
                    if (c == '-' || (c >= '0' && c <= '9'))
                    {
                        newClassName = "_";
                        continue;
                    }
                }

                if ((c >= '0' && c <= '9') ||
                   (c >= 'a' && c <= 'z') ||
                   (c >= 'A' && c <= 'Z') ||
                    c >= 0x00A0)
                {
                    newClassName += c;
                }
            }
            return string.IsNullOrEmpty(newClassName) ? optionalName : newClassName;
        }

        internal static string GetWorksheetClassName(string styleClassPrefix, string name, ExcelWorksheet ws, bool addWorksheetName)
        {
            if (addWorksheetName)
            {
                return styleClassPrefix + name + "-" + GetClassName(ws.Name, $"Sheet{ws.PositionId}");
            }
            else
            {
                return styleClassPrefix + name;
            }
        }

        internal static string GetTableClasses(ExcelTable table)
        {
            string styleClass;
            if (table.TableStyle == TableStyles.Custom)
            {
                styleClass = TableStyleClassPrefix + table.StyleName.Replace(" ", "-").ToLowerInvariant();
            }
            else
            {
                styleClass = TableStyleClassPrefix + table.TableStyle.ToString().ToLowerInvariant();
            }

            var tblClasses = $"{styleClass}";
            if (table.ShowHeader)
            {
                tblClasses += $" {styleClass}-header";
            }

            if (table.ShowTotal)
            {
                tblClasses += $" {styleClass}-total";
            }

            if (table.ShowRowStripes)
            {
                tblClasses += $" {styleClass}-row-stripes";
            }

            if (table.ShowColumnStripes)
            {
                tblClasses += $" {styleClass}-column-stripes";
            }

            if (table.ShowFirstColumn)
            {
                tblClasses += $" {styleClass}-first-column";
            }

            if (table.ShowLastColumn)
            {
                tblClasses += $" {styleClass}-last-column";
            }

            return tblClasses;
        }

        internal static void AddClassesAttributes(EpplusHtmlWriter writer, ExcelTable table, HtmlTableExportSettings settings)
        {
            if (table.TableStyle == TableStyles.None)
            {
                writer.AddAttribute(HtmlAttributes.Class, $"{TableClass}");
            }
            else
            {
                var tblClasses = $"{TableClass} ";
                tblClasses += GetTableClasses(table);
                if (settings.AdditionalTableClassNames.Count > 0)
                {
                    foreach (var cls in settings.AdditionalTableClassNames)
                    {
                        tblClasses += $" {cls}";
                    }
                }

                writer.AddAttribute(HtmlAttributes.Class, tblClasses);
            }
            if (!string.IsNullOrEmpty(settings.TableId))
            {
                writer.AddAttribute(HtmlAttributes.Id, settings.TableId);
            }
        }

        internal static void RenderTableCss(StreamWriter sw, ExcelTable table, HtmlTableExportSettings settings, Dictionary<string, int> styleCache, List<string> datatypes)
        {
            var styleWriter = new EpplusTableCssWriter(sw, table, settings);
            if (settings.Minify == false) styleWriter.WriteLine();
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

            var tableClass = $"{TableClass}.{TableStyleClassPrefix}{GetClassName(tblStyle.Name, "EmptyTableStyle").ToLower()}";
            styleWriter.AddHyperlinkCss($"{tableClass}", tblStyle.WholeTable);
            styleWriter.AddAlignmentToCss($"{tableClass}", datatypes);

            styleWriter.AddToCss($"{tableClass}", tblStyle.WholeTable, "");
            styleWriter.AddToCssBorderVH($"{tableClass}", tblStyle.WholeTable, "");

            //Header
            styleWriter.AddToCss($"{tableClass}", tblStyle.HeaderRow, " thead");
            styleWriter.AddToCssBorderVH($"{tableClass}", tblStyle.HeaderRow, "");

            styleWriter.AddToCss($"{tableClass}", tblStyle.LastTotalCell, $" thead tr th:last-child)");
            styleWriter.AddToCss($"{tableClass}", tblStyle.FirstHeaderCell, " thead tr th:first-child");

            //Total
            styleWriter.AddToCss($"{tableClass}", tblStyle.TotalRow, " tfoot");
            styleWriter.AddToCssBorderVH($"{tableClass}", tblStyle.TotalRow, "");
            styleWriter.AddToCss($"{tableClass}", tblStyle.LastTotalCell, $" tfoot tr td:last-child)");
            styleWriter.AddToCss($"{tableClass}", tblStyle.FirstTotalCell, " tfoot tr td:first-child");

            //Columns stripes
            var tableClassCS = $"{tableClass}-column-stripes";
            styleWriter.AddToCss($"{tableClassCS}", tblStyle.FirstColumnStripe, $" tbody tr td:nth-child(odd)");
            styleWriter.AddToCss($"{tableClassCS}", tblStyle.SecondColumnStripe, $" tbody tr td:nth-child(even)");

            //Row stripes
            var tableClassRS = $"{tableClass}-row-stripes";
            styleWriter.AddToCss($"{tableClassRS}", tblStyle.FirstRowStripe, " tbody tr:nth-child(odd)");
            styleWriter.AddToCss($"{tableClassRS}", tblStyle.SecondRowStripe, " tbody tr:nth-child(even)");

            //Last column
            var tableClassLC = $"{tableClass}-last-column";
            styleWriter.AddToCss($"{tableClassLC}", tblStyle.LastColumn, $" tbody tr td:last-child");

            //First column
            var tableClassFC = $"{tableClass}-first-column";
            styleWriter.AddToCss($"{tableClassFC}", tblStyle.FirstColumn, " tbody tr td:first-child");

            styleWriter.FlushStream();
        }
#if !NET35 && !NET40
        internal static async Task RenderTableCssAsync(StreamWriter sw, ExcelTable table, HtmlTableExportSettings settings, Dictionary<string, int> styleCache, List<string> datatypes)
        {
            var styleWriter = new EpplusTableCssWriter(sw, table, settings);
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
#endif
    }
}
