using EPPlus.DrawingRenderer.RenderItems;
using EPPlus.DrawingRenderer.Svg;
using EPPlus.Export.ImageRenderer.RenderItems.Shared;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlus.DrawingRenderer
{
    public class SvgBasicShapesRenderer : IBasicIShapesRenderer<StringBuilder>
    {
        public SvgBasicShapesRenderer(StringBuilder outputStream)
        {
            LineRenderer = new SvgLineRenderer(outputStream);
            RectangleRenderer = new SvgRectRenderer(outputStream);
            EllipseRenderer = new SvgEllipseRenderer(outputStream);
            PathRenderer = new SvgPathRenderer(outputStream);
            ParagraphRenderer = new SvgParagraphRenderer(this, outputStream);
            GroupRenderer = new SvgGroupRenderer(this, outputStream);
            TextRunRenderer = new SvgTextRunRenderer(outputStream);
            TitleRenderer = new SvgTitleRenderer(outputStream);
            UseReferenceRenderer = new SvgUseReferenceRenderer(outputStream);

            // ImageRenderer = new SvgImageRenderer(outputStream);
        }
        public BaseRenderer<StringBuilder, GroupRenderItem> GroupRenderer { get; }
        public BaseRenderer<StringBuilder, RectRenderItem> RectangleRenderer { get; }
        public BaseRenderer<StringBuilder, EllipseRenderItem> EllipseRenderer { get; }
        public BaseRenderer<StringBuilder, PathRenderItem> PathRenderer { get; }
        //public BaseRenderer<StringBuilder> ImageRenderer { get; }
        public BaseRenderer<StringBuilder, LineRenderItem> LineRenderer { get; }
        public BaseRenderer<StringBuilder, TitleRenderItem> TitleRenderer { get; }
        public BaseRenderer<StringBuilder, ParagraphRenderItem> ParagraphRenderer { get; }
        public BaseRenderer<StringBuilder, TextRunRenderItem> TextRunRenderer { get; }
        public BaseRenderer<StringBuilder, UseReferenceRenderItem> UseReferenceRenderer { get; }

        public void Render(RenderItem item)
        {
            switch (item.Type)
            {
                case RenderItemType.Group:
                    GroupRenderer.Render((GroupRenderItem)item);
                    break;
                case RenderItemType.Line:
                    LineRenderer.Render((LineRenderItem)item);
                    break;
                case RenderItemType.Rect:
                    RectangleRenderer.Render((RectRenderItem)item);
                    break;
                case RenderItemType.Ellipse:
                    EllipseRenderer.Render((EllipseRenderItem)item);
                    break;
                case RenderItemType.Path:
                    PathRenderer.Render((PathRenderItem)item);
                    break;
                case RenderItemType.Text:
                    throw new NotImplementedException();
                    break;
                case RenderItemType.CommentTitle:
                    TitleRenderer.Render((TitleRenderItem)item);
                    break;
                case RenderItemType.Paragraph:
                    ParagraphRenderer.Render((ParagraphRenderItem)item);
                    break;
                case RenderItemType.UseReference:
                    UseReferenceRenderer.Render((UseReferenceRenderItem)item);
                    break;
                case RenderItemType.TextRun:
                    TextRunRenderer.Render((TextRunRenderItem)item);
                    break;
            }
        }
    }
}
