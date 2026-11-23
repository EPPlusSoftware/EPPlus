/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  23/11/2026         EPPlus Software AB           EPPlus 9
 *************************************************************************************************/
using EPPlus.DrawingRenderer.RenderItems;
using EPPlus.DrawingRenderer.Svg;
using EPPlus.Export.ImageRenderer.RenderItems.Shared;
using System.Text;

namespace EPPlus.DrawingRenderer
{
    public class SvgBasicShapesRenderer : IBasicIShapesRenderer<StringBuilder>
    {
        public SvgBasicShapesRenderer(StringBuilder outputStream)
            : this(outputStream, SvgFontPlan.WithoutContext)
        {
        }

        internal SvgBasicShapesRenderer(StringBuilder outputStream, SvgFontPlan fontPlan)
        {
            LineRenderer = new SvgLineRenderer(outputStream);
            RectangleRenderer = new SvgRectRenderer(outputStream);
            EllipseRenderer = new SvgEllipseRenderer(outputStream);
            PathRenderer = new SvgPathRenderer(outputStream);
            ParagraphRenderer = new SvgParagraphRenderer(this, outputStream) { FontPlan = fontPlan };
            GroupRenderer = new SvgGroupRenderer(this, outputStream) { FontPlan = fontPlan };
            TextRunRenderer = new SvgTextRunRenderer(outputStream) { FontPlan = fontPlan };
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
