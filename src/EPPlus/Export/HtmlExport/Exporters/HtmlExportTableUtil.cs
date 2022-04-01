using OfficeOpenXml.Table;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Export.HtmlExport.Exporters
{
    internal static class HtmlExportTableUtil
    {

        internal const string TableStyleClassPrefix = "ts-";
        internal const string TableClass = "epplus-table";

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
    }
}
