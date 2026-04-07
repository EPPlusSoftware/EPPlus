using EPPlus.Fonts.OpenType.Utils;
using EPPlus.Graphics;
using EPPlusImageRenderer;
using EPPlusImageRenderer.RenderItems;
using EPPlusImageRenderer.Svg;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Text;
using OfficeOpenXml.Utils.EnumUtils;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Runtime.Serialization;
using OfficeOpenXml.FormulaParsing.Excel.Functions.MathFunctions;


namespace EPPlus.Export.ImageRenderer.RenderItems.SvgItem
{
    internal class SvgTextBox : DrawingObjectNoBounds
    {
        internal SvgTextBox(DrawingBase renderer, BoundingBox parent, double left, double top, double width, double height, double maxWidth = double.NaN, double maxHeight = double.NaN) : base(renderer)
        {
            Init(renderer, parent, maxWidth, maxHeight);
            Left = left;
            Top = top;
        }

        private void Init(DrawingBase renderer, BoundingBox parent, double maxWidth, double maxHeight)
        {
            Parent = parent;
            _rectangle = new SvgRenderRectItem(DrawingRenderer, Parent);
            TextBody = new SvgTextBodyItem(renderer, Rectangle.Bounds, true);
            TextBody.MaxWidth = maxWidth;
            TextBody.MaxHeight = maxHeight;            
        }

        internal SvgTextBox(DrawingBase renderer, BoundingBox parent, double maxWidth, double maxHeight) : base(renderer)
        {
            Init(renderer, parent, maxWidth, maxHeight);
        }

        //Simplified input
        internal SvgTextBox(DrawingBase renderer, BoundingBox parent, BoundingBox maxBounds) : this(
                                            renderer, parent, maxBounds.Left, maxBounds.Top, maxBounds.Width, maxBounds.Height, maxBounds.Width, maxBounds.Height)
        {
        }
        SvgRenderItem _rectangle=null;
        public SvgRenderItem Rectangle
        {
            get
            {
                _rectangle.Bounds.Width = Width;
                _rectangle.Bounds.Height = Height;
                return _rectangle;
            }
        }
        public SvgTextBodyItem TextBody {get;set;}
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
                 return LeftMargin + (TextBody?.Width ?? 0D) + RightMargin;
            } 
        }
        public double Height
        {
            get
            {
                return TopMargin + (TextBody?.Height ?? 0d) + BottomMargin;
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

        internal void ImportTextBody(ExcelTextBody body, bool useDefaults = true, ExcelHorizontalAlignment horizontalDefault = ExcelHorizontalAlignment.Left)
        {
            double l, r, t, b;
            if (useDefaults)
            {
                body.GetInsetsOrDefaults(out l, out t, out r, out b);
            }
            else
            {
                body.GetInsetsInPoints(out l, out t, out r, out b);
            }
            LeftMargin = l;
            TopMargin = t;
            RightMargin = r;
            BottomMargin = b;

            TextBody.ImportTextBody(body);
        }

        internal override void AppendRenderItems(List<RenderItem> renderItems)
        {
            var rect = Rectangle;

            SvgGroupItem groupItem;
            if (Rotation == 0)
            {
                groupItem = new SvgGroupItem(DrawingRenderer, new BoundingBox(Left, Top, Width, Height));
            }
            else
            {
                groupItem = new SvgGroupItem(DrawingRenderer, new BoundingBox(Left, Top, Width, Height), Rotation);
            }
            groupItem.TextAnchor = TextAnchor.ToEnumString();
            renderItems.Add(groupItem);

            var textboxGroupItem = new SvgGroupItem(DrawingRenderer);
            renderItems.Add(textboxGroupItem);

            var titleItem = new SvgTitleItem(DrawingRenderer, "TextBodySvg Rect");
            //The rect shound encapse the text element, so we need to set the left depending on the text anchor.
            if(TextAnchor==eTextAnchor.Middle)
            {
                rect.Bounds.Left = -(rect.Bounds.Width / 2);
            }
            else if(TextAnchor==eTextAnchor.End)
            {
                rect.Bounds.Left = -rect.Bounds.Width;
            }
            rect.Bounds.Left = 0;
            rect.Bounds.Top = 0;
            renderItems.Add(titleItem);
            renderItems.Add(rect);

            renderItems.Add(new SvgEndGroupItem(DrawingRenderer, rect.Bounds));

            TextBody.Bounds.Left = LeftMargin;
            TextBody.Bounds.Top = TopMargin;
            TextBody.AppendRenderItems(renderItems);
            renderItems.Add(new SvgEndGroupItem(DrawingRenderer, rect.Bounds));
        }

        internal void ImportParagraph(ExcelDrawingParagraph item, double startingY, string text = null)
        {
            TextBody.ImportParagraph(item, startingY, text);
        }

        internal void AddText(double startingY, string text = null)
        {
            TextBody.AddParagraph(startingY, text);
        }
    }
}
