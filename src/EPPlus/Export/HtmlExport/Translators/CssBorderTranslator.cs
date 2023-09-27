using OfficeOpenXml.Drawing.Theme;
using OfficeOpenXml.Export.HtmlExport.Writers.Css;
using OfficeOpenXml.Style;
using OfficeOpenXml.Style.XmlAccess;
using OfficeOpenXml.Utils;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime;
using System.Text;

namespace OfficeOpenXml.Export.HtmlExport.Translators
{
    internal class CssBorderTranslator : TranslatorBase
    {
        ExcelBorderItemXml _top;
        ExcelBorderItemXml _bottom;
        ExcelBorderItemXml _left;
        ExcelBorderItemXml _right;

        ExcelTheme _theme;


        internal CssBorderTranslator(ExcelBorderXml border) 
        {
            _top = border.Top;
            _bottom = border.Bottom;
            _left = border.Left;
            _right = border.Right;
        }

        internal CssBorderTranslator(ExcelBorderItemXml top, ExcelBorderItemXml bottom, ExcelBorderItemXml left, ExcelBorderItemXml right)
        {
            _top = top;
            _bottom = bottom;
            _left = left;
            _right = right;
        }

        internal override List<Declaration> GenerateDeclarationList(TranslatorContext context)
        {
            var borderExclude = context.Exclude.Border;
            _theme = context.Theme;

            if (EnumUtil.HasNotFlag(borderExclude, eBorderExclude.Top)) WriteBorderItem(_top, "top");
            if (EnumUtil.HasNotFlag(borderExclude, eBorderExclude.Bottom)) WriteBorderItem(_bottom, "bottom");
            if (EnumUtil.HasNotFlag(borderExclude, eBorderExclude.Left)) WriteBorderItem(_left, "left");
            if (EnumUtil.HasNotFlag(borderExclude, eBorderExclude.Right)) WriteBorderItem(_right, "right");
            //TODO add Diagonal
            //WriteBorderItem(b.DiagonalDown, "right");
            //WriteBorderItem(b.DiagonalUp, "right");

            return declarations;
        }

        private void WriteBorderItem(ExcelBorderItemXml bi, string suffix)
        {
            if (bi.Style != ExcelBorderStyle.None)
            {
                AddDeclaration($"border-{suffix}", GetBorderItemLine(bi.Style, suffix));

                if (bi.Color != null && bi.Color.Exists)
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
