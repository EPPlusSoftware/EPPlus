using EPPlus.DrawingRenderer.RenderItems;
using EPPlus.Fonts.OpenType.Utils;
using System.Globalization;
using System.Text;

namespace EPPlus.DrawingRenderer.Svg
{
    public class SvgGroupRenderer : SvgBaseRenderer<GroupRenderItem> 
    {
        const string transformTranslate = "translate({0}, {1})";
        const string transformRotate = "rotate({0})";
        const string transformScale = "scale({0}, {1})";

        IBasicIShapesRenderer<StringBuilder> _shapeRenderer;
        public SvgGroupRenderer(IBasicIShapesRenderer<StringBuilder> shapeRenderer, StringBuilder outputStream) : base(outputStream)
        {
            _shapeRenderer = shapeRenderer;
        }
        internal string Suffix = "px";
        public override void Render(GroupRenderItem item)
        {
            string combinedTransform = GetCombinedTransformString(item);

            if (string.IsNullOrEmpty(combinedTransform) == false)
            {
                OutputStream.Append($"<g {GetTransformOrigin(item)} transform=\"{combinedTransform}\" >");
            }
            else
            {
                OutputStream.Append($"<g>");
            }

            foreach (var childItem in item.RenderItems)
            {
                _shapeRenderer.Render(childItem);
            }

            OutputStream.Append("</g>");
        }
        string GetCombinedTransformString(GroupRenderItem item)
        {
            string positionStr = "";
            string rotationStr = GetRotationStr(item);
            string scalingStr = GetScalingStr(item);


            if (item.Position != null && (item.Position.Left == 0 && item.Position.Top == 0) == false)
            {
                positionStr = string.Format(transformTranslate, item.Position.Left.PointToPixelString(), item.Position.Top.PointToPixelString()) + " ";
            }

            return positionStr + rotationStr + scalingStr;
        }
        string GetTransformOrigin(GroupRenderItem item)
        {
            string tOrigin = string.Empty;

            if (item.TransformOrigin != null && (item.TransformOrigin.X == 0 && item.TransformOrigin.Y == 0) == false)
            {
                tOrigin = $"transform-origin=\"{item.TransformOrigin.X.PointToPixelString()} {item.TransformOrigin.Y.PointToPixelString()}\"";
            }
            return tOrigin;
        }


        string GetScalingStr(GroupRenderItem item)
        {
            var scaleStr = string.Empty;

            if (item.Scale != null)
            {
                scaleStr = string.Format(transformScale, item.Scale.X.ToString(CultureInfo.InvariantCulture), item.Scale.Y.ToString(CultureInfo.InvariantCulture)) + " ";
            }

            return scaleStr;
        }

        string GetRotationStr(GroupRenderItem item)
        {
            if (double.IsNaN(item.Rotation) == false || item.Rotation!=0)
            {
                string rot = item.Rotation.ToString(CultureInfo.InvariantCulture);

                if (item.RotationPoint != null && item.RotationPoint != item.Position)
                {
                    rot += $", {item.RotationPoint.Left.PointToPixelString()}, {item.RotationPoint.Top.PointToPixelString()}" + " ";
                }

                return string.Format(transformRotate, rot);
            }

            return string.Empty;
        }

    }
}
