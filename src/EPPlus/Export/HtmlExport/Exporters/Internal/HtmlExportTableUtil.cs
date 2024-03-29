﻿/*************************************************************************************************
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
using OfficeOpenXml.Export.HtmlExport.HtmlCollections;
using OfficeOpenXml.Export.HtmlExport.Settings;
using OfficeOpenXml.Table;
#if !NET35 && !NET40
using System.Threading.Tasks;
#endif

namespace OfficeOpenXml.Export.HtmlExport.Exporters.Internal
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
				   (c == '-') ||
					c >= 0x00A0)
				{
					newClassName += c;
				}
			}
			return string.IsNullOrEmpty(newClassName) ? optionalName.ToLowerInvariant() : newClassName.ToLowerInvariant();
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
                styleClass = TableStyleClassPrefix + GetClassName(table.StyleName, $"tablestyle{table.Id}");
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

        internal static void AddClassesAttributes(HTMLElement element, ExcelTable table, HtmlTableExportSettings settings)
        {
            if (table.TableStyle == TableStyles.None)
            {
                element.AddAttribute(HtmlAttributes.Class, $"{TableClass}");
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

                element.AddAttribute(HtmlAttributes.Class, tblClasses);
            }
            if (!string.IsNullOrEmpty(settings.TableId))
            {
                element.AddAttribute(HtmlAttributes.Id, settings.TableId);
            }
        }
    }
}
