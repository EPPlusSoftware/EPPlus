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
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Export.HtmlExport.CssCollections;
using System.Collections.Generic;
using System.Drawing;

namespace OfficeOpenXml.Export.HtmlExport.Translators
{
    internal class CssImagePropertiesTranslator : TranslatorBase
    {
        double _width;
        double _height;
        ExcelImageInfo _bounds;
        ExcelDrawingBorder _border;

        internal CssImagePropertiesTranslator(HtmlImage image)
        {
            _width = image.Picture.GetPixelWidth();
            _height = image.Picture.GetPixelHeight();
            _bounds = image.Picture.Image.Bounds;
            _border = image.Picture.Border;
        }

        internal override List<Declaration> GenerateDeclarationList(TranslatorContext context)
        {
            if (context.Pictures.KeepOriginalSize == false)
            {
                if (_width != _bounds.Width)
                {
                    AddDeclaration("max-width", $"{_width:F0}px");
                }
                if (_height != _bounds.Height)
                {
                    AddDeclaration("max-height", $"{_height:F0}px");
                }
            }

            if (_border.LineStyle != null && context.Pictures.CssExclude.Border == false)
            {
                var border = GetDrawingBorder();
                AddDeclaration("border", border);
            }

            return declarations;
        }

        private string GetDrawingBorder()
        {
            Color color = _border.Fill.Color;
            if (color.IsEmpty) return "";
            string lineStyle = $"{_border.Width}px";

            switch (_border.LineStyle.Value)
            {
                case eLineStyle.Solid:
                    lineStyle += " solid";
                    break;
                case eLineStyle.Dash:
                case eLineStyle.LongDashDot:
                case eLineStyle.LongDashDotDot:
                case eLineStyle.SystemDash:
                case eLineStyle.SystemDashDot:
                case eLineStyle.SystemDashDotDot:
                    lineStyle += $" dashed";
                    break;
                case eLineStyle.Dot:
                    lineStyle += $" dot";
                    break;
            }

            lineStyle += " #" + color.ToArgb().ToString("x8").Substring(2);
            return lineStyle;
        }
    }
}
