using EPPlus.Export.Utils;
using EPPlus.Graphics;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Export.HtmlExport.CssCollections;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeOpenXml.Export.HtmlExport.Translators
{
    internal class CssDrawingPropertiesTranslator : TranslatorBase
    {

        double _width;
        double _height;
        BoundingBox _bounds;
        ExcelDrawingBorder _border;

        internal CssDrawingPropertiesTranslator(HtmlSvgDrawing d)
        {
            _width = d.Drawing.GetPixelWidth();
            _height = d.Drawing.GetPixelHeight();
            _bounds = d.Drawing.GetBoundingBox();
            if(d.Drawing is ExcelChart)
            {
                _border = d.Drawing.As.Chart.Chart.Border;
            }
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
