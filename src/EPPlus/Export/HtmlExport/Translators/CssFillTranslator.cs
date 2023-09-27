using OfficeOpenXml.Drawing.Theme;
using OfficeOpenXml.Export.HtmlExport.Writers.Css;
using OfficeOpenXml.Style;
using OfficeOpenXml.Style.XmlAccess;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Export.HtmlExport.Translators
{
    internal class CssFillTranslator : TranslatorBase
    {
        ExcelFillXml _fill;
        ExcelTheme _theme;

        internal CssFillTranslator(ExcelFillXml fill) 
        {
            _fill = fill;
        }

        public override List<Declaration> GenerateDeclarationList(TranslatorContext context)
        {
            _theme = context.Theme;

            if (context.FillExclude) return null;
            if (_fill is ExcelGradientFillXml gf && gf.Type != ExcelFillGradientType.None)
            {
                AddGradient(gf);
            }
            else
            {
                if (_fill.PatternType == ExcelFillStyle.Solid)
                {
                    AddDeclaration("background-color", GetColor(_fill.BackgroundColor, _theme));
                }
                else
                {
                    var svg = PatternFills.GetPatternSvgConvertedOnly(_fill.PatternType, GetColor(_fill.BackgroundColor, _theme), GetColor(_fill.PatternColor, _theme));
                    AddDeclaration("background-repeat", "repeat");
                    //arguably some of the values should be its own declaration...Should still work though.
                    AddDeclaration("background", $"url(data:image/svg+xml;base64,{svg})");
                }
            }

            return declarations;
        }

        private void AddGradient(ExcelGradientFillXml gradient)
        {
            AddDeclaration("background");
            var gradientDeclaration = declarations.LastOrDefault();

            if (gradient.Type == ExcelFillGradientType.Linear)
            {
                gradientDeclaration.AddValues($"linear-gradient({(gradient.Degree + 90) % 360}deg");
            }
            else
            {
                gradientDeclaration.AddValues($"radial-gradient(ellipse {gradient.Right * 100}% {gradient.Bottom * 100}%");
            }

            gradientDeclaration.AddValues
                (
                $",{GetColor(gradient.GradientColor1, _theme)} 0%",
                $",{GetColor(gradient.GradientColor2, _theme)} 100%)"
                );
        }
    }
}
