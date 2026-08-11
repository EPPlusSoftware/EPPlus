using DrawingRenderer.Constants;
using EPPlus.DrawingRenderer.RenderItems;
using EPPlus.DrawingRenderer.Svg;
using EPPlus.Export.ImageRenderer.RenderItems.Shared;
using EPPlus.Graphics;
using System.Text;

namespace EPPlus.DrawingRenderer
{
    public interface IBasicIShapesRenderer<T>
    {
        public BaseRenderer<T, GroupRenderItem> GroupRenderer { get; }
        public BaseRenderer<T, RectRenderItem> RectangleRenderer { get; }
        public BaseRenderer<T, EllipseRenderItem> EllipseRenderer { get; }
        public BaseRenderer<T,PathRenderItem> PathRenderer { get; }
        //public BaseRenderer<T> ImageRenderer { get; }
        public BaseRenderer<T,LineRenderItem> LineRenderer { get; }
        public BaseRenderer<T, TitleRenderItem> TitleRenderer { get; }
        public BaseRenderer<T,ParagraphRenderItem> ParagraphRenderer { get; }
        public BaseRenderer<T, UseReferenceRenderItem> UseReferenceRenderer { get; }

        public void Render(RenderItem item);
    }
}
