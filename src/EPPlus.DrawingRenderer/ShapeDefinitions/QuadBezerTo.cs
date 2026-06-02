using System.Xml;

namespace EPPlus.DrawingRenderer.ShapeDefinitions
{
    public class QuadBezerTo : PathWithCoordinates
    {
        public QuadBezerTo(QuadBezerTo clone) : base(clone)
        {

        }
        public QuadBezerTo(XmlReader xr) : base(xr)
        {

        }
        public QuadBezerTo(XmlElement e) : base(e)
        {

        }
        public override PathDrawingType Type => PathDrawingType.QuadBezierTo;
        internal override PathsBase Clone()
        {
            return new QuadBezerTo(this);
        }

    }
}
