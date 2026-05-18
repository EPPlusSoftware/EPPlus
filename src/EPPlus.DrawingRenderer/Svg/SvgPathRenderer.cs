using EPPlus.DrawingRenderer.RenderItems;
using EPPlus.Fonts.OpenType.Utils;
using EPPlusImageRenderer;
using EPPlusImageRenderer.RenderItems;
using System.Globalization;
using System.Text;

namespace EPPlus.DrawingRenderer.Svg
{
    public class SvgPathRenderer : SvgBaseRenderer
    {
        public SvgPathRenderer(StringBuilder outputStream) : base(outputStream)
        {
        }
        public override void Render(RenderItem item)
        {
            var path = (PathRenderItem)item;
            StringBuilder sb = OutputStream;
            //Draw transparent lines to create the compond line effect, as SVG does not support compound lines natively
            switch (path.CompoundLineStyle)
            {
                case CompoundLineStyle.Single:
                    RenderPathItem(path, null, null, null);
                    break;
                case CompoundLineStyle.Double:
                    var name = $"double-stroke-{Guid.NewGuid().ToString()}";
                    sb.Append($"<defs><mask id=\"{name}\">");

                    RenderPathItem(path, path.BorderWidth, "white", null);
                    RenderPathItem(path, path.BorderWidth * (3D / 7D), "black", null);
                    sb.Append($"</mask></defs><rect width=\"100%\" height=\"100%\" fill=\"{path.BorderColor}\" mask=\"url(#{name})\" />");
                    break;
                case CompoundLineStyle.DoubleThickThin:
                    WriteThickThin(path, (path.BorderWidth ?? 1D) * 1D / 7D);
                    break;
                case CompoundLineStyle.DoubleThinThick:
                    WriteThickThin(path, ((path.BorderWidth ?? 1D) * 1D / 7D) * -1);
                    break;
                case CompoundLineStyle.TripleThinThickThin:
                    var guid = Guid.NewGuid().ToString();
                    var gapOffset = 5 * path.BorderWidth.Value / 16;
                    name = $"triple-stroke-{guid}";
                    sb.Append($"<defs>");
                    sb.Append($"<filter id=\"gap-left-{guid}\" x=\"-500%\" y=\"-500%\" width=\"1100%\" height=\"1100%\"><feOffset dx=\"0\" dy=\"-{gapOffset.PointToPixel().ToString(CultureInfo.InvariantCulture)}\" /></filter>");
                    sb.Append($"<filter id=\"gap-right-{guid}\" x=\"-500%\" y=\"-500%\" width=\"1100%\" height=\"1100%\"><feOffset dx=\"0\" dy=\"{gapOffset.PointToPixel().ToString(CultureInfo.InvariantCulture)}\" /></filter>");
                    sb.Append($"<mask id=\"{name}\">");
                    RenderPathItem(path, path.BorderWidth, "white", null);
                    RenderPathItem(path, path.BorderWidth * (1D / 8D), "black", $"filter=\"url(#gap-left-{guid})\"");
                    RenderPathItem(path, path.BorderWidth * (1D / 8D), "black", $"filter=\"url(#gap-right-{guid})\"");
                    sb.Append($"</mask></defs><rect width=\"100%\" height=\"100%\" fill=\"{path.BorderColor}\" mask=\"url(#{name})\" />");
                    break;
            }
        }
        private void RenderPathItem(PathRenderItem path, double? borderWidth, string color, string filter)
        {
            OutputStream.Append($"<path d=\"");
            for (int i = 0; i < path.Commands.Count; i++)
            {
                RenderPathCommand(path.Commands[i]);
            }
            OutputStream.Append("\" ");
            RenderCompoundItems(path, borderWidth, color, filter);
        }

        private void RenderPathCommand(PathCommands pc)
        {
            OutputStream.Append(pc.Type.AsCommandChar());
            for (int i = 0; i < pc.Coordinates.Length; i++)
            {
                string s;
                if (pc.Type == PathCommandType.Arc && ((i & 7) == 2 || (i & 7) == 3 || (i & 7) == 4)) // Arc flags are not coordinates, but should be rendered as integers
                {
                    s = pc.Coordinates[i].ToString(CultureInfo.InvariantCulture);
                }
                else
                {
                    s = pc.Coordinates[i].PointToPixelString();
                }
                OutputStream.AppendFormat("{0} ", s);
            }
            if (pc.Coordinates.Length > 0)
            {
                OutputStream.Remove(OutputStream.Length - 1, 1);
            }
        }

        }

        private void WriteThickThin(PathRenderItem path, double gapOffset)
        {
            var guid = Guid.NewGuid().ToString();
            var name = $"double-thick-thin-stroke-{guid}";
            string gapFilterName = $"f-gap-shift-{guid}";
            OutputStream.Append("<defs>");
            OutputStream.Append($"<filter id=\"{gapFilterName}\" x=\"-50%\" y=\"-50%\" width=\"200%\" height=\"200%\"><feOffset in=\"SourceGraphic\" dy=\"{gapOffset.PointToPixel().ToString(CultureInfo.InvariantCulture)}\"/></filter>");
            OutputStream.Append($"<mask id=\"{name}\">");
            RenderPathItem(path, path.BorderWidth, "white", null);
            RenderPathItem(path, path.BorderWidth * (1 / 4D), "black", $"filter=\"url(#{gapFilterName})\"");
            OutputStream.Append($"</mask></defs><rect width=\"100%\" height=\"100%\" fill=\"{path.BorderColor}\" mask=\"url(#{name})\" />");
        }
    }
