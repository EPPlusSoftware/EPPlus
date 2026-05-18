using EPPlus.DrawingRenderer.RenderItems;
using EPPlus.DrawingRenderer.Utils;
using EPPlus.Graphics;
using EPPlusImageRenderer.RenderItems;


namespace EPPlus.Export.ImageRenderer.RenderItems.SvgItem
{
    public class RenderTextbox
    {
        public RenderTextbox(BoundingBox parent, double left, double top, double width, double height, double maxWidth = double.NaN, double maxHeight = double.NaN) 
        {
            Init(parent, maxWidth, maxHeight);
            Left = left;
            Top = top;
        }

        public void Init(BoundingBox parent, double maxWidth, double maxHeight)
        {
            Parent = parent;
            _rectangle = new RenderRectItem(Parent);
            //TextBody = new TextBody(Rectangle.Bounds, true);
            //TextBody.MaxWidth = maxWidth;
            //TextBody.MaxHeight = maxHeight;            
        }

        internal RenderTextbox(BoundingBox parent, double maxWidth, double maxHeight)
        {
            Init(parent, maxWidth, maxHeight);
        }

        //Simplified input
        internal RenderTextbox(BoundingBox parent, BoundingBox maxBounds) 
        {
        }
        RenderRectItem _rectangle =null;
        public RenderRectItem Rectangle
        {
            get
            {
                _rectangle.Bounds.Width = Width;
                _rectangle.Bounds.Height = Height;
                return _rectangle;
            }
        }
        public virtual RenderTextBody TextBody {get;set;}
        public double Left 
        {
            get
            {
                return Rectangle.Bounds.Left; //TextBody.Bounds.Left - LeftMargin;

            }
            set
            {
                //TextBody.Bounds.Left = value + LeftMargin;
                Rectangle.Bounds.Left = value;
            }
        }
        public double Top 
        { 
            get
            {
                return Rectangle.Bounds.Top;  //TextBody.Bounds.Top - TopMargin;
            }
            set
            {
                //TextBody.Bounds.Top = value + TopMargin;
                Rectangle.Bounds.Top = value;
            } 
        }
        public double Width
        { 
            get 
            {
                 return LeftMargin + (TextBody.Bounds?.Width ?? 0D) + RightMargin;
            } 
        }
        public double Height
        {
            get
            {
                return TopMargin + (TextBody?.Bounds.Height ?? 0d) + BottomMargin;
            }
        }
        internal double LeftMargin
        {
            get; set; 
        }

        internal double TopMargin
        {
            get; set;
        }

        internal double RightMargin
        {
            get; set;
        }

        internal double BottomMargin
        {
            get; set;
        }
        internal BoundingBox Parent { get; private set; }
        internal double Rotation 
        {
            get
            {
                return Rectangle.Bounds.Rotation;
            }
            set
            {
                Rectangle.Bounds.Rotation = value;
            }
        }
        /// <summary>
        /// Gets the actual width of the rotated textbox.
        /// </summary>
        /// <returns></returns>
        internal double GetActualWidth()
        {
            return Width * Math.Abs(Math.Cos(MathHelper.Radians(Rotation))) + Height * Math.Abs(Math.Sin(MathHelper.Radians(Rotation)));
        }
        /// <summary>
        /// Gets the actual right position of the rotated textbox.
        /// </summary>
        /// <returns></returns>
        internal double GetActualRight()
        {
            return Left+GetActualWidth();
        }
        /// <summary>
        /// Gets the actual height of the rotated textbox.
        /// </summary>
        /// <returns></returns>
        internal double GetActualHeight()
        {
            return Width * Math.Abs(Math.Sin(MathHelper.Radians(Rotation))) + Height * Math.Abs(Math.Cos(MathHelper.Radians(Rotation)));
        }
        /// <summary>
        /// Gets the actual right position of the rotated textbox.
        /// </summary>
        /// <returns></returns>
        internal double GetActualBottom()
        {
            return Top + GetActualHeight();
        }
        /// <summary>
        /// How the text is anchored.
        /// </summary>
        internal eTextAnchor TextAnchor
        {
            get;
            set;
        }
    }
}
