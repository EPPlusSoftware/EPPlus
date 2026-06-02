using EPPlus.Export.ImageRenderer.RenderItems;
using EPPlus.Export.ImageRenderer.RenderItems.Interfaces;
using EPPlusImageRenderer;
using EPPlusImageRenderer.RenderItems;
using OfficeOpenXml.Drawing.Style.Fill;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;

namespace EPPlus.Export.ImageRenderer.Svg.DefinitionUtils.UtillNodes
{
    internal class DynamicGridItem : SymbolGroup
    {
        SvgRenderRectItem HorizontalLines = null;
        SvgRenderRectItem VerticalLines = null;

        public DynamicGridItem(DrawingBase renderer, string id, string maskId, string linesHorizontalId, string linesVerticalId) : base(renderer, id)
        {
            Mask = string.Format(_urlRef, maskId);

            HorizontalLines = IntitalizeRect();
            HorizontalLines.FillColor = string.Format(_urlRef, linesHorizontalId);

            VerticalLines = IntitalizeRect();
            VerticalLines.FillColor = string.Format(_urlRef, linesVerticalId);

            _items.Add(HorizontalLines);
            _items.Add(VerticalLines);
        }

        SvgRenderRectItem IntitalizeRect()
        {
            var rect = new SvgRenderRectItem(DrawingRenderer, Bounds);
            rect.Width = 100;
            rect.Height = 100;
            rect.Suffix = "%";

            return rect;
        }
    }
}
