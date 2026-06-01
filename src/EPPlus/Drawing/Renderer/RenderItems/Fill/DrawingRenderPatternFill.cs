using EPPlus.DrawingRenderer.RenderItems;
using OfficeOpenXml.Drawing.Style.Fill;
using OfficeOpenXml.Drawing.Theme;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using tc = OfficeOpenXml.Utils.TypeConversion;

namespace OfficeOpenXml.Drawing.Renderer.RenderItems.Fill
{
    internal class DrawingRenderPatternFill : RenderPatternFill
    {
        private ExcelDrawingPatternFill _patternFill;

        public DrawingRenderPatternFill(ExcelTheme theme, ExcelDrawingPatternFill patternFill, EPPlus.DrawingRenderer.PathFillMode fillColorSource)
        {
            _patternFill = patternFill;
            base.PatternType = (FillPatternStyle)patternFill.PatternType;
            ForegroundColor = tc.ColorConverter.GetThemeColor(theme, _patternFill.ForegroundColor);
            BackgroundColor = tc.ColorConverter.GetThemeColor(theme, _patternFill.BackgroundColor);
        }
    }
}
