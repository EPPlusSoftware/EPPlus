using OfficeOpenXml.Drawing;
using OfficeOpenXml.Export.HtmlExport.Writers.Css;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Text;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Runtime;
using System.Text;

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
