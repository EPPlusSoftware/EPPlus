using EPPlus.DrawingRenderer;

namespace EPPlus.Export.ImageRenderer.RenderItems.SvgItem
{
    public class ConnectionPointsMiddle
    {
        public Coordinate Left;
        public Coordinate Top;
        public Coordinate Right;
        public Coordinate Bottom;

        public Dictionary<int, Coordinate> Points = new Dictionary<int, Coordinate>();

        public ConnectionPointsMiddle(double left, double top, double width, double height)
        {
            var middleWidth = width / 2;
            var middleHeight = height / 2;

            var middleX = left + middleWidth;
            var middleY = top + middleHeight;

            Left = new Coordinate(left, middleY);
            Top = new Coordinate(middleX, top);

            Right = new Coordinate(left + width, middleY);
            Bottom = new Coordinate(left + middleX, top + height);

            Points.Add(0, Left);
            Points.Add(1, Top);
            Points.Add(2, Right);
            Points.Add(3, Bottom);
        }
    }
}
