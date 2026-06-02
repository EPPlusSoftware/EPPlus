using EPPlusImageRenderer;
using EPPlusImageRenderer.RenderItems;
using OfficeOpenXml.FormulaParsing.Excel.Functions.MathFunctions;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;

namespace EPPlus.Export.ImageRenderer.Svg.DefinitionUtils.UtillNodes
{
    internal class FadeOutMask : MaskGroup
    {
        public FadeOutMask(DrawingBase renderer, string id, string rectFillId) : base(renderer, id)
        {
            var rect = new SvgRenderRectItem(DrawingRenderer, Bounds);
            rect.Width = 100;
            rect.Height = 100;
            rect.Suffix = "%";

            rect.FillColor = rectFillId;

            _items.Add(rect);
        }
    }
}
