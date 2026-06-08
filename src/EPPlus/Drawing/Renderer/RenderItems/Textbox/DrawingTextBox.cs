using EPPlus.Graphics;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.FormulaParsing.Excel.Functions.MathFunctions;
using OfficeOpenXml.Style;
using OfficeOpenXml.Utils.EnumUtils;
using System;
using System.Collections.Generic;
using EPPlus.DrawingRenderer.RenderItems;
using EPPlus.Export.ImageRenderer.RenderItems.SvgItem;
namespace OfficeOpenXml.Drawing.Renderer.TextBox
{
    public class DrawingTextBox : RenderTextbox
    {
        ExcelDrawing _drawing;
        internal DrawingTextBox(ExcelDrawing drawing, BoundingBox parent, double left, double top, double width, double height, double maxWidth = double.NaN, double maxHeight = double.NaN) : base(parent, left, top, width, height, maxWidth, maxHeight)
        {
            Init(drawing, parent, maxWidth, maxHeight);
            Left = left;
            Top = top;
        }

        private void Init(ExcelDrawing drawing, BoundingBox parent, double maxWidth, double maxHeight) 
        {
            Parent = parent;
            _drawing= drawing; 
            _rectangle = new RectRenderItem(parent);
            TextBody = new DrawingTextbody(drawing, Rectangle.Bounds, true);
            TextBody.MaxWidth = maxWidth;
            TextBody.MaxHeight = maxHeight;
        }

        internal DrawingTextBox(ExcelDrawing drawing, BoundingBox parent, double maxWidth, double maxHeight) : base(parent, maxWidth, maxHeight)
        {
            Init(drawing, parent, maxWidth, maxHeight);
        }

        ////Simplified input
        //internal DrawingTextBox(BoundingBox parent, BoundingBox maxBounds) : this(
        //                                    parent, maxBounds.Left, maxBounds.Top, maxBounds.Width, maxBounds.Height, maxBounds.Width, maxBounds.Height)
        //{
        //}
        RectRenderItem _rectangle =null;
        public RectRenderItem Rectangle
        {
            get
            {
                _rectangle.Bounds.Width = Width;
                _rectangle.Bounds.Height = Height;
                return _rectangle;
            }
        }
        //internal override void AppendRenderItems(List<RenderItem> renderItems)
        //{
        //    var rect = Rectangle;

        //    SvgGroupItem groupItem;
        //    if (Rotation == 0)
        //    {
        //        groupItem = new SvgGroupItem(DrawingRenderer, new BoundingBox(Left, Top, Width, Height));
        //    }
        //    else
        //    {
        //        groupItem = new SvgGroupItem(DrawingRenderer, new BoundingBox(Left, Top, Width, Height), Rotation);
        //    }
        //    groupItem.TextAnchor = TextAnchor.ToEnumString();
        //    renderItems.Add(groupItem);

        //    var textboxGroupItem = new SvgGroupItem(DrawingRenderer);
        //    renderItems.Add(textboxGroupItem);

        //    var titleItem = new SvgTitleItem(DrawingRenderer, "TextBodySvg Rect");
        //    //The rect shound encapse the text element, so we need to set the left depending on the text anchor.
        //    if (TextAnchor == eTextAnchor.Middle)
        //    {
        //        rect.Bounds.Left = -(rect.Bounds.Width / 2);
        //    }
        //    else if (TextAnchor == eTextAnchor.End)
        //    {
        //        rect.Bounds.Left = -rect.Bounds.Width;
        //    }
        //    else
        //    {
        //        rect.Bounds.Left = 0;
        //    }
        //    rect.Bounds.Top = 0;
        //    renderItems.Add(titleItem);
        //    renderItems.Add(rect);

        //    renderItems.Add(new SvgEndGroupItem(DrawingRenderer, rect.Bounds));

        //    TextBody.Bounds.Left = LeftMargin;
        //    TextBody.Bounds.Top = TopMargin;
        //    TextBody.AppendRenderItems(renderItems);
        //    renderItems.Add(new SvgEndGroupItem(DrawingRenderer, rect.Bounds));
        //}

        internal void AddText(string text = null)
        {
            TextBody.AddParagraph(text);
        }

        public new DrawingTextbody TextBody { get { return (DrawingTextbody)base.TextBody; } set { base.TextBody = value; } }
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

            TextBody.ImportTextBodyAndParagraphs(body, horizontalDefault);
        }

        public override void AppendRenderItems(List<RenderItem> renderItems)
        {
            var rect = Rectangle;

            GroupRenderItem groupItem = groupItem = new GroupRenderItem((BoundingBox)rect.Bounds.Parent, Rotation);
            groupItem.Bounds = new BoundingBox(Left, Top, Width, Height);
            groupItem.TextAnchor = TextAnchor.ToEnumString();
            renderItems.Add(groupItem);

            rect.Top = 0;
            rect.Left = 0;


            var titleItem = new TitleRenderItem("TextBodySvg Rect");
            //The rect shound encapse the text element, so we need to set the left depending on the text anchor.
            if(TextAnchor==eTextAnchor.Middle)
            {
                groupItem.Bounds.Left += -(rect.Bounds.Width / 2);
            }
            else if(TextAnchor==eTextAnchor.End)
            {
                if (Math.Abs(Rotation) == 45)
                {
                    const double COS45 = 0.70710678118654757; //Constant for Math.Sin(Math.PI / 4) --45 degrees
                    groupItem.Bounds.Left += -(rect.Bounds.Width * COS45);
                    groupItem.Bounds.Top += (rect.Bounds.Width * COS45);
                }
                else
                {
                    groupItem.Bounds.Left += rect.Bounds.Height / 2;
                    groupItem.Bounds.Top += (rect.Bounds.Width);
                }
            }
            groupItem.RenderItems.Add(titleItem);

            if(TextBody.AutoSize)
            {
                TextBody.ApplyAutoSize();
                rect.Width = TextBody.Width + RightMargin + LeftMargin;
                rect.Height = TextBody.Height + BottomMargin + RightMargin;
            }

            //As the rect item is inside the group, we set the left and right to the group and top and left on the rect to 0.
            groupItem.RenderItems.Add(rect);

            TextBody.Bounds.Left = LeftMargin;
            TextBody.Bounds.Top = TopMargin;
            TextBody.AppendRenderItems(groupItem.RenderItems);
        }

        internal void ImportParagraph(ExcelDrawingParagraph item, double startingY, string text = null)
        {
            TextBody.ImportParagraph(item, startingY, text);
        }

        //internal void AddText(double startingY, string text = null)
        //{
        //    TextBody.AddParagraph(startingY, text);
        //}
    }
}
