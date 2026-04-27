using EPPlus.Export.ImageRenderer.RenderItems.Shared;
using EPPlus.Fonts.OpenType.Utils;
using EPPlus.Graphics;
using EPPlusImageRenderer;
using System.Globalization;
using System.Text;

namespace EPPlus.Export.ImageRenderer.RenderItems.SvgItem
{
    internal class SvgGroupItemNew : GroupItem
    {
        const string transformTranslate = "translate({0}, {1})";
        const string transformRotate = "rotate({0})";
        const string transformScale = "scale({0}, {1})";

        public SvgGroupItemNew(DrawingBase renderer, double localXPos, double localYPos) : base(renderer, localXPos, localYPos)
        {
        }

        public SvgGroupItemNew(DrawingBase renderer, BoundingBox parent, double rotation, Transform rotationPoint = null) : base(renderer, parent, rotation, rotationPoint)
        {
        }

        public override void Render(StringBuilder sb)
        {
            string combinedTransform = GetCombinedTransformString();

            if (string.IsNullOrEmpty(combinedTransform) == false)
            {
                sb.Append($"<g {GetTransformOrigin()} transform=\"{combinedTransform}\" >");
            }
            else
            {
                sb.Append($"<g>");
            }

            foreach (var item in _childItems)
            {
                item.Render(sb);
            }

            sb.Append("</g>");
        }

        string GetTransformOrigin()
        {
            string tOrigin = string.Empty;

            if(TransformOrigin != null && (TransformOrigin.X == 0 && TransformOrigin.Y == 0) == false)
            {
                tOrigin = $"transform-origin=\"{TransformOrigin.X.PointToPixelString()} {TransformOrigin.Y.PointToPixelString()}\"";
            }
            return tOrigin;
        }

        string GetCombinedTransformString()
        {
            string positionStr = "";
            string rotationStr = GetRotationStr();
            string scalingStr = GetScalingStr();


            if (Position != null && (Position.Left == 0 && Position.Top == 0) == false)
            {
                positionStr = string.Format(transformTranslate, Position.Left.PointToPixelString(), Position.Top.PointToPixelString()) + " ";
            }
            
            return positionStr + rotationStr + scalingStr;
        }

        string GetScalingStr()
        {
            var scaleStr = string.Empty;

            if (Scale != null)
            {
                scaleStr = string.Format(transformScale, Scale.X.ToString(CultureInfo.InvariantCulture), Scale.Y.ToString(CultureInfo.InvariantCulture)) + " ";
            }

            return scaleStr;
        }

        string GetRotationStr()
        {
            if (double.IsNaN(Rotation) == false)
            {
                string rot = Rotation.ToString(CultureInfo.InvariantCulture);

                if(RotationPoint != null && RotationPoint != Position)
                {
                    rot += $", {RotationPoint.Left.PointToPixelString()}, {RotationPoint.Top.PointToPixelString()}" + " ";
                }

                return string.Format(transformRotate, rot);
            }

            return string.Empty;
        }
    }
}
