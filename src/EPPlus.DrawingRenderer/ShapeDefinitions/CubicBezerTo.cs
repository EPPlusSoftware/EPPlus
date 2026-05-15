using System.Xml;

namespace EPPlus.DrawingRenderer.ShapeDefinitions
{
    public class CubicBezerTo : PathWithCoordinates
    {
        public CubicBezerTo(CubicBezerTo clone) : base(clone)
        {

        }
        public CubicBezerTo(XmlReader xr) : base(xr)
        {

        }
        public CubicBezerTo(XmlElement e) : base(e)
        {

        }

        public override PathDrawingType Type => PathDrawingType.CubicBezierTo;
        internal override PathsBase Clone()
        {
            return new CubicBezerTo(this);
        }
    }
}
