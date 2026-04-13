using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Theme;
using OfficeOpenXml.Export.HtmlExport;
using OfficeOpenXml.Export.HtmlExport.CssCollections;
using OfficeOpenXml.Export.HtmlExport.Translators;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
namespace EPPlus.Export.ImageRenderer.Style
{
    internal class SvgFillTranslator : TranslatorBase
    {
        ExcelTheme _theme;
        IFillBasic _fill;

        internal SvgFillTranslator(IFillBasic fill)
        {
            _fill = fill;
        }

        internal override List<Declaration> GenerateDeclarationList(TranslatorContext context)
        {
            _theme = context.Theme;

            if (context.Exclude.Fill) return null;

            if (_fill.Style == eFillStyle.GradientFill)
            {
                AddGradient();
            }
            else
            {
                if (_fill.Style == eFillStyle.PatternFill)
                {
                    var bc = _fill.GetBackgroundColor(_theme) ?? "#0";
                    if (string.IsNullOrEmpty(bc) == false)
                    {
                        AddDeclaration("fill", bc);
                    }
                }
                else if (_fill.Style == eFillStyle.NoFill)
                {
                    var fc = _fill.GetBackgroundColor(_theme);
                    if (string.IsNullOrEmpty(fc) == false)
                    {
                        AddDeclaration("fill", fc);
                    }
                }
                else if(_fill.Style == eFillStyle.BlipFill)
                {
                    string bgColor = _fill.GetBackgroundColor(_theme) ?? "#0";
                    //string patternColor = _fill.GetPatternColor(_theme) ?? "#0";

                    var svg = PatternFills.GetPatternSvgConvertedOnly(ExcelFillStyle.Solid, bgColor, "");
                    AddDeclaration("background-repeat", "repeat");
                    //arguably some of the values should be its own declaration...Should still work though.
                    AddDeclaration("background", $"url(data:image/svg+xml;base64,{svg})");
                }
            }

            return declarations;
        }

        private void AddGradient()
        {
            //AddDeclaration("linearGradient");

            //GetLinearGradientNodeStop
            //AddDeclaration("stop");
            //var gradientDeclaration = declarations.LastOrDefault();

            //if (_fill.Style == eFillStyle.GradientFill)
            //{
            //    AddDeclaration("offset",)
            //}
            //else
            //{
            //    gradientDeclaration.AddValues($"radial-gradient(ellipse {_fill.Right * 100}% {_fill.Bottom * 100}%");
            //}

            //gradientDeclaration.AddValues
            //    (
            //    $",{_fill.GetGradientColor1(_theme)} 0%",
            //    $",{_fill.GetGradientColor2(_theme)} 100%)"
            //    );
        }
    }
}
