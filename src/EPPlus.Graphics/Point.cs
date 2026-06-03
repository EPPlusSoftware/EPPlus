using EPPlus.Graphics.Geometry;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlus.Graphics
{
    public class Point : Transform
    {
        public double Top
        {
            get { return LocalPosition.Y; }
            set
            {
                LocalPosition = new Vector2(LocalPosition.X, value);
            }
        }

        public double Left
        {
            get { return LocalPosition.X; }
            set
            {
                LocalPosition = new Vector2(value, LocalPosition.Y);
            }
        }

        public Point()
        {
            
        }

        public Point(double x, double y)
        {
            Left = x;
            Top = y;
        }
    }
}
