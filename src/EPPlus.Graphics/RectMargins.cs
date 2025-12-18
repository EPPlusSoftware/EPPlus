using OfficeOpenXml.Drawing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlus.Graphics
{
    internal class RectMargins : Rect
    {
        internal double MarginLeft { get; set; }
        internal double MarginTop { get; set; }
        internal double MarginRight { get; set; }
        internal double MarginBottom { get; set; }

        internal RectMargins() : base() { }

        internal RectMargins(double l, double t, double r, double b) : base(l, t, r, b)
        {
        }

        internal double GetInnerLeft()
        {
            return Left + MarginLeft;
        }

        internal double GetInnerTop()
        {
            return Top + MarginTop;
        }

        internal double GetInnerBottom()
        {
            return Bottom - MarginBottom;
        }

        internal double GetInnerRight()
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
