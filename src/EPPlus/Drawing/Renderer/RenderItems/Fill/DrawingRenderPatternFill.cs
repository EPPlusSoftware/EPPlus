using EPPlus.DrawingRenderer.RenderItems;
using OfficeOpenXml.Drawing.Style.Fill;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeOpenXml.Drawing.Renderer.RenderItems.Fill
{
    internal class DrawingRenderPatternFill : RenderPatternFill
    {
        private ExcelDrawingPatternFill _patternFill;

        public DrawingRenderPatternFill(ExcelDrawingPatternFill patternFill)
        {
            _patternFill = patternFill;
        }
    }
}
