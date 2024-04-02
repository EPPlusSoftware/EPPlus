/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  03/14/2024         EPPlus Software AB           Epplus 7.1
 *************************************************************************************************/

using OfficeOpenXml.Drawing.Theme;
using OfficeOpenXml.Export.HtmlExport.CssCollections;
using OfficeOpenXml.Export.HtmlExport.StyleCollectors.StyleContracts;
using OfficeOpenXml.Style;
using System.Collections.Generic;
using System.Linq;

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
					var bc = _fill.GetBackgroundColor(_theme) ?? "#0";
					if (string.IsNullOrEmpty(bc) == false)
                    {
                        AddDeclaration("background-color", bc);
                    }
                }
                else if(_fill.PatternType == ExcelFillStyle.None)
                {
                    var fc = _fill.GetPatternColor(_theme);
                    if (string.IsNullOrEmpty(fc) == false)
                    {
                        AddDeclaration("background-color", fc);
                    }
				}
                else
                {
					string bgColor = _fill.GetBackgroundColor(_theme)??"#0";
					string patternColor = _fill.GetPatternColor(_theme)??"#0";

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
