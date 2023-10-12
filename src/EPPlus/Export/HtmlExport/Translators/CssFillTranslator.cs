using OfficeOpenXml.Drawing.Theme;
using OfficeOpenXml.Export.HtmlExport.StyleCollectors;
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
        //ExcelFillXml _fill = null;
        //ExcelDxfFill _dxfFill = null;
        ExcelTheme _theme;
        //double _degree;
        //double _right;
        //double _bottom;
        //bool _isLinear = false;
        //bool _isSolid = false;
        //bool _isGradient = false;
        GenericFill _fill;


        internal CssFillTranslator(GenericFill fill)
        {
            _fill = fill;
        }

        internal override List<Declaration> GenerateDeclarationList(TranslatorContext context)
        {
            _theme = context.Theme;

            if (context.Exclude.Fill) return null;

            if (_fill._isGradient)
            {
                AddGradient();
            }
            else
            {
                if (_fill._isSolid)
                {
                    AddDeclaration("background-color", _fill._color1.GetHexCodeColor(_theme));
                }
                else
                {
                    string bgColor = _fill._color1.GetHexCodeColor(_theme);
                    string patternColor = _fill._color2.GetHexCodeColor(_theme);

                    var svg = PatternFills.GetPatternSvgConvertedOnly(_fill._patternType, bgColor, patternColor);
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

            if (_fill._isLinear)
            {
                gradientDeclaration.AddValues($"linear-gradient({(_fill._degree + 90) % 360}deg");
            }
            else
            {
                gradientDeclaration.AddValues($"radial-gradient(ellipse {_fill._right * 100}% {_fill._bottom * 100}%");
            }

            gradientDeclaration.AddValues
                (
                $",{_fill._color1.GetHexCodeColor(_theme)} 0%",
                $",{_fill._color1.GetHexCodeColor(_theme)} 100%)"
                );
        }


        //internal CssFillTranslator(ExcelFillXml fill)
        //{
        //    _fill = fill;
        //    _isSolid = _fill.PatternType == ExcelFillStyle.Solid;
        //    if (_fill is ExcelGradientFillXml gf && gf.Type != ExcelFillGradientType.None)
        //    {
        //        _isGradient = true;

        //        _degree = gf.Degree;
        //        _right = gf.Right;
        //        _bottom = gf.Bottom;
        //        _isLinear = gf.Type == ExcelFillGradientType.Linear;
        //    }
        //}

        //internal CssFillTranslator(ExcelDxfFill fill)
        //{
        //    _dxfFill = fill;
        //    _isSolid = _dxfFill.PatternType == ExcelFillStyle.Solid;
        //    if(_dxfFill.Gradient.HasValue)
        //    {
        //        _isGradient = true;

        //        _degree = _dxfFill.Gradient.Degree.Value;
        //        _right = _dxfFill.Gradient.Right.Value;
        //        _bottom = _dxfFill.Gradient.Bottom.Value;
        //        _isLinear = _dxfFill.Gradient.GradientType == eDxfGradientFillType.Linear;
        //    }
        //}

        //internal override List<Declaration> GenerateDeclarationList(TranslatorContext context)
        //{
        //    _theme = context.Theme;

        //    if (context.Exclude.Fill) return null;

        //    if (_isGradient)
        //    {
        //        string col1 = "";
        //        string col2 = "";

        //        if (_fill is ExcelGradientFillXml gf)
        //        {
        //            col1 = GetColor(gf.GradientColor1, _theme);
        //            col2 = GetColor(gf.GradientColor2, _theme);
        //        }
        //        else
        //        {
        //            col1 = GetColor(_dxfFill.Gradient.Colors[0].Color, _theme);
        //            col2 = GetColor(_dxfFill.Gradient.Colors[1].Color, _theme);
        //        }

        //        AddGradient(col1, col2);
        //    }
        //    else
        //    {
        //        string bgColor = "";
        //        if (_fill != null)
        //        {
        //            bgColor = GetColor(_fill.BackgroundColor, _theme);
        //        }
        //        else
        //        {
        //            bgColor = GetColor(_dxfFill.BackgroundColor, _theme);
        //        }

        //        if (_isSolid)
        //        {
        //            AddDeclaration("background-color", bgColor);
        //        }
        //        else
        //        {
        //            string patternColor = "";
        //            if (_fill != null)
        //            {
        //                patternColor = GetColor(_fill.PatternColor, _theme);
        //            }
        //            else
        //            {
        //                patternColor = GetColor(_dxfFill.PatternColor, _theme);
        //            }

        //            var svg = PatternFills.GetPatternSvgConvertedOnly(_fill.PatternType, bgColor, patternColor);
        //            AddDeclaration("background-repeat", "repeat");
        //            //arguably some of the values should be its own declaration...Should still work though.
        //            AddDeclaration("background", $"url(data:image/svg+xml;base64,{svg})");
        //        }
        //    }

        //    return declarations;
        //}

        //private void AddGradient(string color1, string color2)
        //{
        //    AddDeclaration("background");
        //    var gradientDeclaration = declarations.LastOrDefault();

        //    if (_isLinear)
        //    {
        //        gradientDeclaration.AddValues($"linear-gradient({(_degree + 90) % 360}deg");
        //    }
        //    else
        //    {
        //        gradientDeclaration.AddValues($"radial-gradient(ellipse {_right * 100}% {_bottom * 100}%");
        //    }

        //    gradientDeclaration.AddValues
        //        (
        //        $",{color1} 0%",
        //        $",{color2} 100%)"
        //        );
        //}
    }
}
