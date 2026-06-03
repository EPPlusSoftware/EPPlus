namespace EPPlus.Graphics
{
    public class RectBase
    {
        public double Left { get; set; }
        internal double Top { get; set; }
        public double Right { get; set; }
        public double Bottom { get; set; }

        public RectBase()
        {
        }

        public RectBase(double width, double height)
        {
            Left = 0;
            Top = 0;
            Right = width;
            Bottom = height;
        }
        public RectBase(double left, double top, double right, double bottom)
        {
            Left = left;
            Top = top;
            Right = right;
            Bottom = bottom;
        }

        public double Width
        {
            get
            {
                return Right - Left;
            }
            set
            {
                Right = Left + value;
            }
        }

        public double Height
        {
            get
            {
                return Bottom - Top;
            }
            set
            {
                Bottom = Top + value;
            }
        }
    }
}
