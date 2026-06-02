using OfficeOpenXml.Drawing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlus.Graphics
{
    public class RectMargins : Rect
    {
        public double MarginLeft { get; set; }
        public double MarginTop { get; set; }
        public double MarginRight { get; set; }
        public double MarginBottom { get; set; }

        public RectMargins() : base() { }

        public RectMargins(double l, double t, double r, double b) : base(l, t, r, b)
        {
        }

        public double GetInnerLeft()
        {
            return Left + MarginLeft;
        }

        public double GetInnerTop()
        {
            return Top + MarginTop;
        }

        public double GetInnerBottom()
        {
            return Bottom - MarginBottom;
        }

        public double GetInnerRight()
        {
            return Right - MarginRight;
        }

        public RectBase GetInnerRect()
        {
            var textArea = new RectBase();
            textArea.Left = GetInnerLeft();
            textArea.Top = GetInnerTop();

            textArea.Bottom = GetInnerBottom();
            textArea.Right = GetInnerRight();

            return textArea;
        }

        public double GetInnerWidth()
        {
            return GetInnerRight() - GetInnerLeft();
        }

        public double GetInnerHeight()
        {
            return GetInnerBottom() - GetInnerTop();
        }
    }
}
