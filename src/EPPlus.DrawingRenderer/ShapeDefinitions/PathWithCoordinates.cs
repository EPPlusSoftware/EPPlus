using EPPlus.DrawingRenderer.Utils;
using System.Globalization;
using System.Xml;

namespace EPPlus.DrawingRenderer.ShapeDefinitions
{
    public abstract class PathWithCoordinates : PathsBase
    {
        protected PathWithCoordinates(XmlElement e)
        {
            foreach (var cn in e.ChildNodes)
            {
                if (cn is XmlElement ce && ce.LocalName == "pt")
                {
                    Coordinates.Add(new DrawCoordinate(GetNameOrNumber(ce.GetAttribute("x")), GetNameOrNumber(ce.GetAttribute("y"))));
                }
            }
        }

        private object GetNameOrNumber(string s)
        {
            if (long.TryParse(s, NumberStyles.Number, CultureInfo.InvariantCulture, out var l))
            {
                return l;
            }
            return s;
        }

        protected PathWithCoordinates(XmlReader xr)
        {
            var name = xr.LocalName;
            while (xr.Read())
            {
                if (xr.LocalName == "pt" && xr.NodeType == XmlNodeType.Element)
                {
                    Coordinates.Add(new DrawCoordinate(GetNameOrNumber(xr.GetAttribute("x")), GetNameOrNumber(xr.GetAttribute("y"))));
                }
                else if (xr.IsEndElementWithName(name))
                {
                    break;
                }
            }
        }

        protected PathWithCoordinates(PathWithCoordinates clone)
        {
            foreach (var c in clone.Coordinates)
            {
                Coordinates.Add(new DrawCoordinate(c));
            }
        }
        public List<DrawCoordinate> Coordinates { get; set; } = new List<DrawCoordinate>();
        public override void TranslateCoordiantesToPointsAndDegrees(double coordinateRatio, double angleRatio)
        {
            foreach (var c in Coordinates)
            {
                c.X /= coordinateRatio;
                c.Y /= coordinateRatio;
            }
        }
        public override double EndX => Coordinates.Count > 0D ? Coordinates[Coordinates.Count - 1].X.Value : 0D;
        public override double EndY => Coordinates.Count > 0D ? Coordinates[Coordinates.Count - 1].Y.Value : 0D;
    }
}
