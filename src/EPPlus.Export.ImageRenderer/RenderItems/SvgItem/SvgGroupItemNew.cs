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
        const string transformTranslate = "translate({0}, {1}) ";
        const string transformRotate = "rotate({0}) ";

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
                sb.Append($"<g transform=\"{combinedTransform}\">");
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

        string GetCombinedTransformString()
        {
            string positionStr = "";
            string rotationStr = GetRotationStr();

            if (Position != null)
            {
                positionStr = string.Format(transformTranslate, Position.Top.PointToPixelString(), Position.Left.PointToPixelString()) + " ";
            }
            
            return positionStr + rotationStr;
        }

        string GetRotationStr()
        {
            if (double.IsNaN(Rotation) == false)
            {
                string rot = Rotation.ToString(CultureInfo.InvariantCulture);

                if(RotationPoint != null && RotationPoint != Position)
                {
                    rot += $", {RotationPoint.Left.PointToPixelString()}, {RotationPoint.Top.PointToPixelString()}";
                }

                return string.Format(transformRotate, rot);
            }

            return string.Empty;
        }
    }
}
