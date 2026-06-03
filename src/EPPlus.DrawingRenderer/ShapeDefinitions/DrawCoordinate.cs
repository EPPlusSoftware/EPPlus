namespace EPPlus.DrawingRenderer.ShapeDefinitions
{
    public class DrawCoordinate
    {
        public DrawCoordinate(DrawCoordinate c)
        {
            X = c.X;
            Y = c.Y;
            XName = c.XName;
            YName = c.YName;
        }

        public DrawCoordinate(object x, object y)
        {
            if (x is long xl)
            {
                X = xl;
            }
            else
            {
                XName = x.ToString();
                X = null;
            }
            if (y is long yl)
            {
                Y = yl;
            }
            else
            {
                YName = y.ToString();
                Y = null;
            }

        }
        public double? X { get; set; }
        public double? Y { get; set; }
        public string XName { get; set; }
        public string YName { get; set; }
    }
}
