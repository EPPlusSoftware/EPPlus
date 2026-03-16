using EPPlusImageRenderer;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Text;
using System.Collections.Generic;


namespace EPPlus.Export.ImageRenderer.RenderItems.SvgItem
{
    internal class ConnectionPointsMiddle
    {
        internal Coordinate Left;
        internal Coordinate Top;
        internal Coordinate Right;
        internal Coordinate Bottom;

        internal Dictionary<int, Coordinate> Points;

        internal ConnectionPointsMiddle(double left, double top, double width, double height)
        {
            var middleWidth = width / 2;
            var middleHeight = height / 2;

            Left = new Coordinate(left, middleHeight);
            Top= new Coordinate(middleWidth, top);

            Right = new Coordinate(left + width, middleHeight);
            Bottom = new Coordinate(middleWidth, top + height);

            Points = new Dictionary<int, Coordinate>();

            Points.Add(0, Left);
            Points.Add(1, Top);
            Points.Add(2, Right);
            Points.Add(3, Bottom);
        }
    }
}
