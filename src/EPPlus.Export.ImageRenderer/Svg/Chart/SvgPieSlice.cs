using EPPlus.Export.ImageRenderer.RenderItems.SvgItem;
using EPPlus.Export.ImageRenderer.Utils;
using EPPlus.Graphics;
using EPPlusImageRenderer;
using EPPlusImageRenderer.RenderItems;
using EPPlusImageRenderer.Svg;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlus.Export.ImageRenderer.Svg.Chart
{
    internal class SvgPieSlice : SvgRenderItem
    {
        double _radius;
        SvgGroupItemNew _innerGroup;
        double _sliceDegrees;
        Point _circleMidPoint;


        public SvgPieSlice(DrawingBase renderer, BoundingBox parent, Point circleMidPoint, double radius, double sliceDegrees, double prevSliceDegrees) : base(renderer, parent)
        {
            _sliceDegrees = sliceDegrees;
            _radius = radius;
            _innerGroup = new SvgGroupItemNew(renderer, parent, 0, parent);
            _circleMidPoint = circleMidPoint;
        }

        Point CalculateLocalPointOnCircle(double degrees)
        {
            var angleRadians = MConverter.DegreesToRadians(degrees);

            //This is already offset by cx and cy because the circle midpoint is the parent
            var xPoint = _radius * Math.Cos(angleRadians);
            var yPoint = _radius * Math.Sin(angleRadians);

            var point = new Point(xPoint, yPoint);

            point.Parent = _circleMidPoint;

            return point;
        }

        public override RenderItemType Type => RenderItemType.Group;

        internal override SvgRenderItem Clone(SvgShape svgDocument)
        {
            throw new NotImplementedException();
        }
    }
}
