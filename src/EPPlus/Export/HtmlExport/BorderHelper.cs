using OfficeOpenXml.Style;
using System.IO;
namespace OfficeOpenXml.Export.HtmlExport
{
    internal static class BorderHelper
    {
        internal static string WriteBorderItemLine(ExcelBorderStyle style, string suffix)
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

    }
}
