namespace OfficeOpenXml.Drawing
{
    internal class RectBase
    {
        internal double Left { get; set; }
        internal double Top { get; set; }
        internal double Right { get; set; }
        internal double Bottom { get; set; }

        internal RectBase()
        {
        }

        internal RectBase(double width, double height)
        {
            Left = 0;
            Top = 0;
            Right = width;
            Bottom = height;
        }
        internal RectBase(double left, double top, double right, double bottom)
        {
            Left = left;
            Top = top;
            Right = right;
            Bottom = bottom;
        }

        internal double Width
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

        internal double Height
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
