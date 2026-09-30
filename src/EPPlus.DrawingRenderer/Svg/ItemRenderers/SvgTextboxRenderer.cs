using EPPlus.Export.ImageRenderer.RenderItems.Shared;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlus.DrawingRenderer.Svg
{
    public class SvgTextboxRenderer : SvgBaseRenderer<TextboxRenderItem>
    {
        IBasicIShapesRenderer<StringBuilder> _shapeRenderer;
        public SvgTextboxRenderer(IBasicIShapesRenderer<StringBuilder> shapeRenderer, StringBuilder outputStream) : base(outputStream)
        {
            _shapeRenderer = shapeRenderer;
        }

        public override void Render(TextboxRenderItem textbox)
        {
            if (textbox.HasBackground) _shapeRenderer.RectangleRenderer.Render(textbox);
            foreach(var p in textbox.Paragraphs)
            {
                _shapeRenderer.ParagraphRenderer.Render(p);
            }
        }
    }
}
