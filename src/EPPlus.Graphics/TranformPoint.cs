using EPPlus.Graphics.Geometry;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlus.Graphics
{
    public class TranformPoint : Transform
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

        public TranformPoint()
        {
            
        }

        public TranformPoint(double x, double y)
        {
            Left = x;
            Top = y;
        }
    }
}
