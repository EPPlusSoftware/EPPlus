using OfficeOpenXml.Drawing.Theme;
using OfficeOpenXml.Export.HtmlExport.StyleCollectors;
using OfficeOpenXml.Export.HtmlExport.StyleCollectors.StyleContracts;
using OfficeOpenXml.Export.HtmlExport.Writers.Css;
using OfficeOpenXml.FormulaParsing.Excel.Functions.MathFunctions;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Text;
using OfficeOpenXml.Style;
using OfficeOpenXml.Style.Dxf;
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
        ExcelTheme _theme;
        IFill _fill;

        internal CssFillTranslator(IFill fill)
        {
            _fill = fill;
        }

        internal override List<Declaration> GenerateDeclarationList(TranslatorContext context)
        {
            _theme = context.Theme;

            if (context.Exclude.Fill) return null;

            if (_fill.IsGradient)
            {
                AddGradient();
            }
            else
            {
                if (_fill.PatternType == ExcelFillStyle.Solid)
                {
                    AddDeclaration("background-color", _fill.GetBackgroundColor(_theme));
                }
                else
                {
                    string bgColor = _fill.GetBackgroundColor(_theme);
                    string patternColor = _fill.GetPatternColor(_theme);

                    var svg = PatternFills.GetPatternSvgConvertedOnly(_fill.PatternType, bgColor, patternColor);
                    AddDeclaration("background-repeat", "repeat");
                    //arguably some of the values should be its own declaration...Should still work though.
                    AddDeclaration("background", $"url(data:image/svg+xml;base64,{svg})");
                }
            }

            return declarations;
        }

        private void AddGradient()
        {
            AddDeclaration("background");
            var gradientDeclaration = declarations.LastOrDefault();

            if (_fill.IsLinear)
            {
                gradientDeclaration.AddValues($"linear-gradient({(_fill.Degree + 90) % 360}deg");
            }
            else
            {
                gradientDeclaration.AddValues($"radial-gradient(ellipse {_fill.Right * 100}% {_fill.Bottom * 100}%");
            }

            gradientDeclaration.AddValues
                (
                $",{_fill.GetGradientColor1(_theme)} 0%",
                $",{_fill.GetGradientColor2(_theme)} 100%)"
                );
        }
    }
}
