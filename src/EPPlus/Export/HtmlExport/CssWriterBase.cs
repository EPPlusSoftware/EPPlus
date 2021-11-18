/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  11/07/2021         EPPlus Software AB       Added Html Export
 *************************************************************************************************/
using OfficeOpenXml.Style;
using OfficeOpenXml.Style.XmlAccess;
using static OfficeOpenXml.Export.HtmlExport.ColumnDataTypeManager;

namespace OfficeOpenXml.Export.HtmlExport
{
    internal abstract class CssWriterBase
    {
        protected CssTableExportOptions _options;
        internal int Indent { get; set; }

        protected static string WriteBorderItemLine(ExcelBorderStyle style, string suffix)
        {
            var lineStyle = $"border-{suffix}:";
            switch (style)
            {
                case ExcelBorderStyle.Hair:
                    lineStyle += "1px solid";
                    break;
                case ExcelBorderStyle.Thin:
                    lineStyle += $"thin solid";
                    break;
                case ExcelBorderStyle.Medium:
                    lineStyle += $"medium solid";
                    break;
                case ExcelBorderStyle.Thick:
                    lineStyle += $"thick solid";
                    break;
                case ExcelBorderStyle.Double:
                    lineStyle += $"double";
                    break;
                case ExcelBorderStyle.Dotted:
                    lineStyle += $"dotted";
                    break;
                case ExcelBorderStyle.Dashed:
                case ExcelBorderStyle.DashDot:
                case ExcelBorderStyle.DashDotDot:
                    lineStyle += $"dashed";
                    break;
                case ExcelBorderStyle.MediumDashed:
                case ExcelBorderStyle.MediumDashDot:
                case ExcelBorderStyle.MediumDashDotDot:
                    lineStyle += $"medium dashed";
                    break;
            }
            return lineStyle;
        }
        protected static string GetVerticalAlignment(ExcelXfs xfs)
        {
            switch (xfs.VerticalAlignment)
            {
                case ExcelVerticalAlignment.Top:
                    return "top";
                case ExcelVerticalAlignment.Center:
                    return "middle";
                case ExcelVerticalAlignment.Bottom:
                    return "bottom";
            }

            return "";
        }

        protected static string GetHorizontalAlignment(ExcelXfs xfs)
        {
            switch (xfs.HorizontalAlignment)
            {
                case ExcelHorizontalAlignment.Right:
                    return "right";
                case ExcelHorizontalAlignment.Center:
                case ExcelHorizontalAlignment.CenterContinuous:
                    return "center";
                case ExcelHorizontalAlignment.Left:
                    return "left";
            }

            return "";
        }

    }
}