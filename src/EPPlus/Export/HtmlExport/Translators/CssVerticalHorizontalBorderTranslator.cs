using OfficeOpenXml.Export.HtmlExport.CssCollections;
using System.Collections.Generic;
using System.Linq;
using OfficeOpenXml.Style.Dxf;
using OfficeOpenXml.Drawing.Theme;
using OfficeOpenXml.Style;
using OfficeOpenXml.Utils;

namespace OfficeOpenXml.Export.HtmlExport.Translators
{
    internal class CssVerticalHorizontalBorderTranslator : TranslatorBase
    {
        ExcelDxfBorderBase _border;
        ExcelTheme _theme;

        public CssVerticalHorizontalBorderTranslator(ExcelDxfBorderBase border)
        {
            _border = border;
        }

        internal override List<Declaration> GenerateDeclarationList(TranslatorContext context)
        {
            if (_border.HasValue)
            {
                var borderExclude = context.Exclude.Border;
                _theme = context.Theme;

                if (EnumUtil.HasNotFlag(borderExclude, eBorderExclude.Top)) WriteBorderItem(_border.Horizontal, "top");
                if (EnumUtil.HasNotFlag(borderExclude, eBorderExclude.Bottom)) WriteBorderItem(_border.Horizontal, "bottom");
                if (EnumUtil.HasNotFlag(borderExclude, eBorderExclude.Left)) WriteBorderItem(_border.Vertical, "left");
                if (EnumUtil.HasNotFlag(borderExclude, eBorderExclude.Right)) WriteBorderItem(_border.Vertical, "right");
            }

            return declarations;
        }

        private void WriteBorderItem(ExcelDxfBorderItem bi, string suffix)
        {
            if (bi.HasValue && bi.Style != ExcelBorderStyle.None)
            {
                AddDeclaration($"border-{suffix}", GetBorderItemLine(bi.Style.Value, suffix));

                if (bi.Color.HasValue)
                {
                    declarations.Last().AddValues(GetColor(bi.Color, _theme));
                }
            }
        }

        protected static string GetBorderItemLine(ExcelBorderStyle style, string suffix)
        {
            var lineStyle = "";
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
