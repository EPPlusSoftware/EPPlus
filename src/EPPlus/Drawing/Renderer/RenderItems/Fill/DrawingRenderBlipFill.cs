using EPPlus.DrawingRenderer.RenderItems;
using OfficeOpenXml.Drawing.Style.Fill;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeOpenXml.Drawing.Renderer.RenderItems.Fill
{
    public class DrawingRenderBlipFill : RenderBlipFill
    {
        private ExcelDrawingBlipFill _blipFill;

        internal DrawingRenderBlipFill(ExcelDrawingBlipFill blipFill)
        {
            _blipFill = blipFill;
        }
    }
}
