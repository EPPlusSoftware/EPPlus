using EPPlus.DrawingRenderer.RenderItems;
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
            if (item.Bounds.Name != null)
            {
                sb.Append($" id=\"{item.Bounds.Name}\" ");
            }

            if (string.IsNullOrEmpty(item.DefId) == false)
            {
                sb.Append($"id=\"{item.DefId}\" ");
            }

            if (string.IsNullOrEmpty(item.FillColor) == false)
            {
                sb.Append($"fill=\"{item.FillColor}\" ");
            }
            //If fill is null it may in e.g. Rect still get the color black which can have an opacity
            if (item.FillOpacity != null && item.FillOpacity != 1)
            {
                sb.Append($"opacity=\"{item.FillOpacity.Value.ToString(CultureInfo.InvariantCulture)}\" ");
            }
            if (string.IsNullOrEmpty(item.FilterName) == false)
            {
                sb.Append($"filter=\"{item.FilterName}\" ");
            }

            if (item.BorderWidth.HasValue)
            {
                if (string.IsNullOrEmpty(item.BorderColor) == false)
                {
                    sb.Append($"stroke=\"{item.BorderColor}\" ");
                }
                var v = item.BorderWidth.Value * Constants.EMU_PER_POINT / Constants.EMU_PER_PIXEL;
                sb.Append($"stroke-width=\"{v.ToString(CultureInfo.InvariantCulture)}\" ");

                if (item.BorderDashArray != null)
                {
                    var BorderDashArrayStr = item.BorderDashArray.Select(x =>
                    x.ToString(CultureInfo.InvariantCulture)).ToArray();

                    sb.Append($"stroke-dasharray=\"" + $"{string.Join(",", BorderDashArrayStr)}\" ");
                }
                if (item.BorderOpacity.HasValue)
                {
                    sb.Append($" stroke-opacity=\"{(Math.Round(item.BorderOpacity.Value * 100)).ToString(CultureInfo.InvariantCulture)}%\" ");
                }
            }

            if (item.TransformOrigin != null)
            {
                sb.Append($" transform-origin=\"{item.TransformOrigin.X.ToString(CultureInfo.InvariantCulture)} {item.TransformOrigin.Y.ToString(CultureInfo.InvariantCulture)}\" ");
            }

            if (item.StrokeMiterLimit.HasValue)
            {
                sb.Append($"stroke-miterlimit =\"{item.StrokeMiterLimit}\" ");
            }
        }

        protected void RenderBase(T item)
        {
            var sb = OutputStream;
            RenderBaseToSpecified(item, sb);
        }
        protected void RenderCompoundItems(T li, double? borderWidth, string color, string filter)
        {
            var tmpBorderWidth = li.BorderWidth;
            string tmpBorderColor = null;
            li.BorderWidth = borderWidth ?? li.BorderWidth;
            if (string.IsNullOrEmpty(color) == false)
            {
                tmpBorderColor = li.BorderColor;
                li.BorderColor = color;
            }

            RenderBase(li);
            var sb = OutputStream;
            if (li.LineCap != LineCap.Flat)
            {
                sb.AppendFormat(" stroke-linecap=\"{0}\"", li.LineCap == LineCap.Round ? "round" : "square");
            }
            if (li.LineJoin != LineJoin.Miter)
            {
                sb.AppendFormat(" stroke-linejoin=\"{0}\"", li.LineJoin);
            }

            if (string.IsNullOrEmpty(filter) == false)
            {
                sb.Append(" " + filter);
            }

            sb.AppendFormat("/>");

            li.BorderWidth = tmpBorderWidth;
            if (string.IsNullOrEmpty(color) == false)
            {
                li.BorderColor = tmpBorderColor;
            }
        }

    }
}
