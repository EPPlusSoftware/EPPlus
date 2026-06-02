using System.Xml;

namespace EPPlus.DrawingRenderer.ShapeDefinitions
{
    public class LineTo : PathWithCoordinates
    {
        public LineTo(LineTo clone) : base(clone)
        {

        }
        public LineTo(XmlReader xr) : base(xr)
        {

        }

        public LineTo(XmlElement e) : base(e)
        {

        }
        public override PathDrawingType Type => PathDrawingType.LineTo;
        public DrawCoordinate Coordinate { get; set; }
        internal override PathsBase Clone()
        {
            return new LineTo(this);
        }
    }
}
