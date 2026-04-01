using EPPlus.Graphics.Math;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlus.Graphics
{
    internal class Point : Transform
    {
        internal double Top
        {
            get { return LocalPosition.Y; }
            set
            {
                LocalPosition = new Vector2(LocalPosition.X, value);
            }
        }

        internal double Left
        {
            get { return LocalPosition.X; }
            set
            {
                LocalPosition = new Vector2(value, LocalPosition.Y);
            }
        }

        internal Point()
        {

        }

        internal Point(double x, double y)
        {
            Left = x;
            Top = y;
        }
    }
}
