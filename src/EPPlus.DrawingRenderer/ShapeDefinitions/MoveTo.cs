using System.Xml;

namespace EPPlus.DrawingRenderer.ShapeDefinitions
{
    public class MoveTo : PathWithCoordinates
    {
        public MoveTo(MoveTo clone) : base(clone)
        {

        }
        public MoveTo(XmlElement e) : base(e)
        {
        }
        public MoveTo(XmlReader xr) : base(xr)
        {
        }
        public override PathDrawingType Type => PathDrawingType.MoveTo;
        public DrawCoordinate Coordinate { get; set; }

        internal override PathsBase Clone()
        {
            return new MoveTo(this);
        }
    }
}
