using OfficeOpenXml.Drawing.Theme;
using OfficeOpenXml.Export.HtmlExport.Writers.Css;
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
        ExcelFillXml _fill = null;
        ExcelDxfFill _dxfFill = null;
        ExcelTheme _theme;
        double _degree;
        double _right;
        double _bottom;
        bool _isLinear = false;

        bool _isSolid = false;

        internal CssFillTranslator(ExcelFillXml fill)
        {
            _fill = fill;
            _isSolid = _fill.PatternType == ExcelFillStyle.Solid;
            if (_fill is ExcelGradientFillXml gf && gf.Type != ExcelFillGradientType.None)
            {
                _degree = gf.Degree;
                _right = gf.Right;
                _bottom = gf.Bottom;
                _isLinear = gf.Type == ExcelFillGradientType.Linear;
            }
        }

        internal CssFillTranslator(ExcelDxfFill fill)
        {
            _dxfFill = fill;
            _isSolid = _dxfFill.PatternType == ExcelFillStyle.Solid;
            if(_dxfFill.Gradient.HasValue)
            {
                _degree = _dxfFill.Gradient.Degree.Value;
                _right = _dxfFill.Gradient.Right.Value;
                _bottom = _dxfFill.Gradient.Bottom.Value;
                _isLinear = _dxfFill.Gradient.GradientType == eDxfGradientFillType.Linear;
            }
        }

        internal override List<Declaration> GenerateDeclarationList(TranslatorContext context)
        {
            _theme = context.Theme;

            if (context.Exclude.Fill) return null;

            if (_fill is ExcelGradientFillXml gf && gf.Type != ExcelFillGradientType.None)
            {
                AddGradient(gf);
            }
            else
            {
                string bgColor = "";
                if (_fill != null)
                {
                    bgColor = GetColor(_fill.BackgroundColor, _theme);
                }
                else
                {
                    bgColor = GetColor(_dxfFill.BackgroundColor, _theme);
                }

                if (_isSolid)
                {
                    AddDeclaration("background-color", bgColor);
                }
                else
                {
                    string patternColor = "";
                    if (_fill != null)
                    {
                        patternColor = GetColor(_fill.PatternColor, _theme);
                    }
                    else
                    {
                        patternColor = GetColor(_dxfFill.PatternColor, _theme);
                    }

                    var svg = PatternFills.GetPatternSvgConvertedOnly(_fill.PatternType, bgColor, patternColor);
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

            if (_isLinear)
            {
                gradientDeclaration.AddValues($"linear-gradient({(_degree + 90) % 360}deg");
            }
            else
            {
                gradientDeclaration.AddValues($"radial-gradient(ellipse {_right * 100}% {_bottom * 100}%");
            }

            gradientDeclaration.AddValues
                (
                $",{GetColor(gradient.GradientColor1, _theme)} 0%",
                $",{GetColor(gradient.GradientColor2, _theme)} 100%)"
                );
        }
    }
}
