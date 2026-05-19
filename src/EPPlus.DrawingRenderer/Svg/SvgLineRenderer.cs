using EPPlus.DrawingRenderer.RenderItems;
using EPPlus.Fonts.OpenType.Utils;
using System.Globalization;
using System.Text;

namespace EPPlus.DrawingRenderer.Svg
{
    public class SvgLineRenderer : SvgBaseRenderer
    {
        public SvgLineRenderer(StringBuilder outputStream) : base(outputStream)
        {
        }
        public override void Render(RenderItem item)
        {
            var li = (LineRenderItem)item;
            StringBuilder sb = OutputStream;
            //Draw transparent lines to create the compond line effect, as SVG does not support compound lines natively
            switch (li.CompoundLineStyle)
            {
                case CompoundLineStyle.Double:
                    li.LineCap = LineCap.Flat;
                    var name = $"double-stroke-{Guid.NewGuid().ToString()}";
                    sb.Append($"<defs><mask id=\"{name}\">");

                    RenderLineItem(li, li.BorderWidth, "white", null);
                    RenderLineItem(li, li.BorderWidth * (3D / 7D), "black", null);
                    sb.Append($"</mask></defs><rect width=\"100%\" height=\"100%\" fill=\"{li.BorderColor}\" mask=\"url(#{name})\" />");
                    break;
                case CompoundLineStyle.DoubleThickThin:
                    WriteThickThin(li, "double-thick-thin-stroke-{0}", (li.BorderWidth ?? 1D) * 1D / 7D);
                    break;
                case CompoundLineStyle.DoubleThinThick:
                    WriteThickThin(li, "double-thin-thick-stroke-{0}", ((li.BorderWidth ?? 1D) * 1D / 7D) * -1);
                    break;
                case CompoundLineStyle.TripleThinThickThin:
                    var guid = Guid.NewGuid().ToString();
                    var gapOffset = 5 * li.BorderWidth.Value / 16;
                    name = $"triple-stroke-{guid}";
                    sb.Append($"<defs>");
                    sb.Append($"<filter id=\"gap-left-{guid}\" x=\"-500%\" y=\"-500%\" width=\"1100%\" height=\"1100%\" filterUnits=\"userSpaceOnUse\"><feOffset dx=\"0\" dy=\"-{gapOffset.PointToPixel().ToString(CultureInfo.InvariantCulture)}\" /></filter>");
                    sb.Append($"<filter id=\"gap-right-{guid}\" x=\"-500%\" y=\"-500%\" width=\"1100%\" height=\"1100%\" filterUnits=\"userSpaceOnUse\"><feOffset dx=\"0\" dy=\"{gapOffset.PointToPixel().ToString(CultureInfo.InvariantCulture)}\" /></filter>");
                    sb.Append($"<mask id=\"{name}\">");
                    RenderLineItem(li, li.BorderWidth, "white", null);
                    RenderLineItem(li, li.BorderWidth * (1D / 8D), "black", $"filter=\"url(#gap-left-{guid})\"");
                    RenderLineItem(li, li.BorderWidth * (1D / 8D), "black", $"filter=\"url(#gap-right-{guid})\"");
                    sb.Append($"</mask></defs><rect width=\"100%\" height=\"100%\" fill=\"{li.BorderColor}\" mask=\"url(#{name})\" />");
                    break;
                default:
                    RenderLineItem(li, null, null, null);
                    break;
            }
        }
        private void WriteThickThin(LineRenderItem li, string name, double gapOffset)
        {
            var sb = OutputStream;
            var guid = Guid.NewGuid().ToString();
            name = string.Format(name, guid);
            string gapFilterName = $"f-gap-shift-{guid}";
            sb.Append("<defs>");
            sb.Append($"<filter id=\"{gapFilterName}\" x=\"-50%\" y=\"-50%\" width=\"200%\" height=\"200%\" filterUnits=\"userSpaceOnUse\"><feOffset in=\"SourceGraphic\" dy=\"{gapOffset.PointToPixel().ToString(CultureInfo.InvariantCulture)}\"/></filter>");
            sb.Append($"<mask id=\"{name}\">");
            RenderLineItem(li, li.BorderWidth, "white", null);
            RenderLineItem(li, li.BorderWidth * (1D / 4D), "black", $"filter=\"url(#{gapFilterName})\"");
            sb.Append($"</mask></defs><rect width=\"100%\" height=\"100%\" fill=\"{li.BorderColor}\" mask=\"url(#{name})\" />");
        }
        internal string Suffix = "px";

        private void RenderLineItem(LineRenderItem li, double? borderWidth, string color, string filter)
        {
            var sb = OutputStream;
            if (Suffix == "%")
            {
                
                sb.AppendFormat("<line x1=\"{0}\" y1=\"{1}\" x2=\"{2}\" y2=\"{3}\" ",
                li.X1.ToString(CultureInfo.InvariantCulture) + Suffix,
                li.Y1.ToString(CultureInfo.InvariantCulture) + Suffix,
                li.X2.ToString(CultureInfo.InvariantCulture) + Suffix,
                li.Y2.ToString(CultureInfo.InvariantCulture) + Suffix);
            }
            else
            {
                sb.AppendFormat("<line x1=\"{0}\" y1=\"{1}\" x2=\"{2}\" y2=\"{3}\" ",
                li.X1.PointToPixelString(),
                li.Y1.PointToPixelString(),
                li.X2.PointToPixelString(),
                li.Y2.PointToPixelString());
            }

            RenderCompoundItems(li, li.BorderWidth, color, filter);
        }
    }
}
