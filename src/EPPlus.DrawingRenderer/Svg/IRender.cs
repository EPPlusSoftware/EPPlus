using EPPlus.DrawingRenderer.RenderItems;
using EPPlus.DrawingRenderer.Svg;
using EPPlus.Fonts.OpenType.Utils;
using EPPlus.Graphics;
using EPPlusImageRenderer;
using EPPlusImageRenderer.RenderItems;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlus.DrawingRenderer.Svg
{
    public class SvgBasicItemsRenderer : IBasicItemsRenderer<StringBuilder>
    {
        public SvgBasicItemsRenderer(StringBuilder outputStream)
        {
            LineRenderer = new SvgLineRenderer(outputStream);
            RectangleRenderer = new SvgRectRenderer(outputStream);
            EllipseRenderer = new SvgEllipseRenderer(outputStream);
            PathRenderer = new SvgPathRenderer(outputStream);
            PathRenderer = new SvgTextRenderer(outputStream);
            // ImageRenderer = new SvgImageRenderer(outputStream);
        }
        public BaseRenderer<StringBuilder> RectangleRenderer { get; }
        public BaseRenderer<StringBuilder> EllipseRenderer { get; }
        public BaseRenderer<StringBuilder> PathRenderer { get; }
        //public BaseRenderer<StringBuilder> ImageRenderer { get; }
        public BaseRenderer<StringBuilder> LineRenderer { get; }
        public BaseRenderer<StringBuilder> TextRenderer { get; }
    }
    public interface IBasicItemsRenderer<T>
    {        
        public BaseRenderer<T> RectangleRenderer { get; }
        public BaseRenderer<T> EllipseRenderer { get; }
        public BaseRenderer<T> PathRenderer { get; }
        //public BaseRenderer<T> ImageRenderer { get; }
        public BaseRenderer<T> LineRenderer { get; }
    }
    public abstract class BaseRenderer<T>
    {
        protected BaseRenderer(T outputStream)
        {
            OutputStream = outputStream;
        }
        public T OutputStream { get; }
        public abstract void Render(RenderItem item);
    }
}
