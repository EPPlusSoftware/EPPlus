using EPPlus.DrawingRenderer.RenderItems;
using EPPlus.DrawingRenderer.Utils;
using System.Globalization;
using System.Text;

namespace EPPlus.DrawingRenderer.Svg
{
    public abstract class SvgBaseRenderer<T> : BaseRenderer<StringBuilder,T> where T : RenderItem
    {
        protected SvgBaseRenderer(StringBuilder outputStream) : base(outputStream)
        {
            
        }

        /// <summary>
        /// Used if you wish to render base to a different string builder first
        /// </summary>
        /// <param name="item"></param>
        /// <param name="sb"></param>
        protected void RenderBaseToSpecified(T item, StringBuilder sb)
        {
            var style = item.Style;
            if (item.Name != null)
            {
                sb.Append($" class=\"{item.Name}\" ");
            }

            if (string.IsNullOrEmpty(style.DefId) == false)
            {
                sb.Append($"id=\"{style.DefId}\" ");
            }

            if (string.IsNullOrEmpty(style.FillColor) == false)
            {
                sb.Append($"fill=\"{style.FillColor}\" ");
            }
            //If fill is null it may in e.g. Rect still get the color black which can have an opacity
            if (style.FillOpacity != null && style.FillOpacity != 1)
            {
                sb.Append($"opacity=\"{style.FillOpacity.Value.ToString(CultureInfo.InvariantCulture)}\" ");
            }
            if (string.IsNullOrEmpty(style.FilterName) == false)
            {
                sb.Append($"filter=\"{style.FilterName}\" ");
            }

            if (style.BorderWidth.HasValue)
            {
                if (string.IsNullOrEmpty(style.BorderColor) == false)
                {
                    sb.Append($"stroke=\"{style.BorderColor}\" ");
                }
                var v = style.BorderWidth.Value * Constants.EMU_PER_POINT / Constants.EMU_PER_PIXEL;
                sb.Append($"stroke-width=\"{v.ToString(CultureInfo.InvariantCulture)}\" ");

                if (style.BorderDashArray != null)
                {
                    var BorderDashArrayStr = style.BorderDashArray.Select(x =>
                    x.ToString(CultureInfo.InvariantCulture)).ToArray();

                    sb.Append($"stroke-dasharray=\"" + $"{string.Join(",", BorderDashArrayStr)}\" ");
                }
                if (style.BorderOpacity.HasValue)
                {
                    sb.Append($" stroke-opacity=\"{(Math.Round(style.BorderOpacity.Value * 100)).ToString(CultureInfo.InvariantCulture)}%\" ");
                }
            }

            if (item.TransformOrigin != null)
            {
                sb.Append($" transform-origin=\"{item.TransformOrigin.X.ToString(CultureInfo.InvariantCulture)} {item.TransformOrigin.Y.ToString(CultureInfo.InvariantCulture)}\" ");
            }

            if (style.StrokeMiterLimit.HasValue)
            {
                sb.Append($"stroke-miterlimit =\"{style.StrokeMiterLimit}\" ");
            }
        }

        protected void RenderBase(T item)
        {
            var sb = OutputStream;
            RenderBaseToSpecified(item, sb);
        }
        protected void RenderCompoundItems(T li, double? borderWidth, string color, string filter)
        {
            var style = li.Style;
            var tmpBorderWidth = style.BorderWidth;
            string tmpBorderColor = null;
            style.BorderWidth = borderWidth ?? style.BorderWidth;
            if (string.IsNullOrEmpty(color) == false)
            {
                tmpBorderColor = style.BorderColor;
                style.BorderColor = color;
            }

            RenderBase(li);
            var sb = OutputStream;
            if (style.LineCap != LineCap.Flat)
            {
                sb.AppendFormat(" stroke-linecap=\"{0}\"", style.LineCap == LineCap.Round ? "round" : "square");
            }
            if (style.LineJoin != LineJoin.Miter)
            {
                sb.AppendFormat(" stroke-linejoin=\"{0}\"", style.LineJoin.ToEnumString());
            }

            if (string.IsNullOrEmpty(filter) == false)
            {
                sb.Append(" " + filter);
            }

            sb.AppendFormat("/>");

            style.BorderWidth = tmpBorderWidth;
            if (string.IsNullOrEmpty(color) == false)
            {
                style.BorderColor = tmpBorderColor;
            }
        }
    }
}
