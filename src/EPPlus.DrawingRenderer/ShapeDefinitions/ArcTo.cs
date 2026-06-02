using System.Xml;

namespace EPPlus.DrawingRenderer.ShapeDefinitions
{
    public class ArcTo : PathsBase
    {
        public ArcTo(XmlReader xr)
        {
            if (long.TryParse(xr.GetAttribute("hR"), out var hrv))
            {
                HeightRadius = hrv;
            }
            else
            {
                HeightRadiusName = xr.GetAttribute("hR");
            }

            if (long.TryParse(xr.GetAttribute("wR"), out var wrv))
            {
                WidthRadius = wrv;
            }
            else
            {
                WidthRadiusName = xr.GetAttribute("wR");
            }

            if (long.TryParse(xr.GetAttribute("swAng"), out var swAng))
            {
                SwingAngle = swAng;
            }
            else
            {
                SwingAngleName = xr.GetAttribute("swAng");
            }

            if (long.TryParse(xr.GetAttribute("stAng"), out var stAng))
            {
                StartAngle = stAng;
            }
            else
            {
                StartAngleName = xr.GetAttribute("stAng");
            }
        }
        public ArcTo(XmlElement e)
        {
            if (long.TryParse(e.GetAttribute("hR"), out var hrv))
            {
                HeightRadius = hrv;
            }
            else
            {
                HeightRadiusName = e.GetAttribute("hR");
            }

            if (long.TryParse(e.GetAttribute("wR"), out var wrv))
            {
                WidthRadius = wrv;
            }
            else
            {
                WidthRadiusName = e.GetAttribute("wR");
            }

            if (long.TryParse(e.GetAttribute("swAng"), out var swAng))
            {
                SwingAngle = swAng;
            }
            else
            {
                SwingAngleName = e.GetAttribute("swAng");
            }

            if (long.TryParse(e.GetAttribute("stAng"), out var stAng))
            {
                StartAngle = stAng;
            }
            else
            {
                StartAngleName = e.GetAttribute("stAng");
            }
        }
        public override PathDrawingType Type => PathDrawingType.ArcTo;
        public double? HeightRadius { get; set; }
        public double? StartAngle { get; set; }
        public double? SwingAngle { get; set; }
        public double? WidthRadius { get; set; }
        public string HeightRadiusName { get; set; }
        public string StartAngleName { get; set; }
        public string SwingAngleName { get; set; }
        public string WidthRadiusName { get; set; }
        private ArcTo()
        {

        }
        internal override PathsBase Clone()
        {
            return new ArcTo()
            {
                HeightRadius = HeightRadius,
                StartAngle = StartAngle,
                SwingAngle = SwingAngle,
                WidthRadius = WidthRadius,
                HeightRadiusName = HeightRadiusName,
                StartAngleName = StartAngleName,
                SwingAngleName = SwingAngleName,
                WidthRadiusName = WidthRadiusName
            };
        }
        double _endX, _endY;
        public void SetEndCoordinates(double x, double y)
        {
            _endX = x;
            _endY = y;
        }
        public override void TranslateCoordiantesToPointsAndDegrees(double coordinateRatio, double angleRatio)
        {
            HeightRadius /= coordinateRatio;
            WidthRadius /= coordinateRatio;
            StartAngle /= angleRatio;
            SwingAngle /= angleRatio;
            _endX = _endX / coordinateRatio;
            _endY = _endY / coordinateRatio;
        }
        public override double EndX => _endX;
        public override double EndY => _endY;
    }
}
