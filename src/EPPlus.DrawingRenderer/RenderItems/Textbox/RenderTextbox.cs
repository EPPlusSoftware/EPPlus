using EPPlus.DrawingRenderer.RenderItems;
using EPPlus.DrawingRenderer.Utils;
using EPPlus.Graphics;
namespace EPPlus.Export.ImageRenderer.RenderItems.SvgItem
{
    public class RenderTextbox : DrawingObject
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
            _group = new GroupRenderItem(Parent);
            _group.Name = "Textbox_Group";
            _rectangle = new RectRenderItem(_group);
            _marginGroup = new GroupRenderItem(_group);
            //TextBody = new RenderTextBody(Rectangle.Bounds, true);
            //TextBody.MaxWidth = maxWidth;
            //TextBody.MaxHeight = maxHeight;
        }

        public RenderTextbox(BoundingBox parent, double maxWidth, double maxHeight)
        {
            Init(parent, maxWidth, maxHeight);
        }

        public string Name 
        { 
            get { return _group.Name; } 
            set { _group.Name = value; } 
        }

        //The origin point of the entire textbox itself (its outermost left and top point)
        protected GroupRenderItem _group;
        //The origin point of the textbody after applied margins
        protected GroupRenderItem _marginGroup;


        protected RectRenderItem _rectangle =null;
        public RectRenderItem Rectangle
        {
            get
            {
                _rectangle.Width = Width;
                _rectangle.Height = Height;
                return _rectangle;
            }
            set
            {
                _rectangle = value;
            }
        }

        RenderTextBody _textBody;

        public virtual RenderTextBody TextBody 
        { 
            get { return _textBody; } 
            set 
            {   _textBody = value; 
                //Margins should affect textbody global position in real-time
                _textBody.Parent = _marginGroup; 
            } 
        }
        public double Left 
        {
            get
            {
                return _group.Left;

            }
            set
            {
                _group.Left = value;
            }
        }
        public double Top 
        { 
            get
            {
                return _group.Top;
            }
            set
            {
                _group.Top = value;
            } 
        }
        public double Width
        { 
            get 
            {
                 return LeftMargin + (TextBody?.Width ?? 0D) + RightMargin + TextBody.Left;
            } 
        }
        public double WidthRotated
        {
            get
            {
                var radians = MathHelper.Radians(Rotation);
                var sin = Math.Abs(Math.Sin(radians));
                var cos = Math.Abs(Math.Cos(radians));
                return Width * sin + Height * cos;
            }
        }
        public double Height
        {
            get
            {
                return TopMargin + (TextBody?.Height ?? 0d) + BottomMargin + TextBody.Top;
            }
        }
        public double HeightWithRotation 
        {
            get
            {
                var radians = MathHelper.Radians(Rotation);
                var sin = Math.Abs(Math.Sin(radians));
                var cos = Math.Abs(Math.Cos(radians));
                return Width * cos + Height * sin;
            }
        }
        public double LeftMargin
        {
            get { return _marginGroup.Left; } set { _marginGroup.Left = value; }
        }

        public double TopMargin
        {
            get { return _marginGroup.Top; }
            set { _marginGroup.Top = value; }
        }

        public double RightMargin
        {
            get; set;
        }

        public double BottomMargin
        {
            get; set;
        }

        internal protected BoundingBox Parent { get; protected set; }
        public double Rotation 
        {
            get
            {
                return _group.Rotation;
            }
            set
            {
                _group.Rotation = value;
            }
        }
        /// <summary>
        /// Gets the actual width of the rotated textbox.
        /// </summary>
        /// <returns></returns>
        public double GetActualWidth()
        {
            return Width * Math.Abs(Math.Cos(MathHelper.Radians(Rotation))) + Height * Math.Abs(Math.Sin(MathHelper.Radians(Rotation)));
        }
        /// <summary>
        /// Gets the actual right position of the rotated textbox.
        /// </summary>
        /// <returns></returns>
        public double GetActualRight()
        {
            return Left+GetActualWidth();
        }
        /// <summary>
        /// Gets the actual height of the rotated textbox.
        /// </summary>
        /// <returns></returns>
        public double GetActualHeight()
        {
            return Width * Math.Abs(Math.Sin(MathHelper.Radians(Rotation))) + Height * Math.Abs(Math.Cos(MathHelper.Radians(Rotation)));
        }
        /// <summary>
        /// Gets the actual right position of the rotated textbox.
        /// </summary>
        /// <returns></returns>
        public double GetActualBottom()
        {
            return Top + GetActualHeight();
        }

        public void PrepareForRender()
        {
            //As the rect item is inside the group, we set the left and right to the group and top and left on the rect to 0.
            _group.Left = Left;
            _group.Top = Top;
            _group.Width = Width;
            _group.Height = Height;

            _group.TextAnchor = TextAnchor.ToEnumString();
            Rectangle.Top = 0;
            Rectangle.Left = 0;

            if (TextBody.AutoSize)
            {
                TextBody.ApplyAutoSize();
            }

            Rectangle.Width = Width;
            Rectangle.Height = Height;

            var titleItem = new TitleRenderItem("TextBox group");
            _group.ChildObjects.Add(titleItem);
            //The rect shound encapse the text element, so we need to set the left depending on the text anchor.
            if (TextAnchor == eTextAnchor.Middle)
            {
                _group.Left += -(Rectangle.Width / 2);
            }
            else if (TextAnchor == eTextAnchor.End)
            {
                if (Math.Abs(Rotation) == 45)
                {
                    const double COS45 = 0.70710678118654757; //Constant for Math.Sin(Math.PI / 4) --45 degrees
                    _group.Left += -(Rectangle.Width * COS45);
                    _group.Top += (Rectangle.Width * COS45);
                }
                else
                {
                    _group.Left += Rectangle.Height / 2;
                    _group.Top += (Rectangle.Width);
                }
            }

            //The textbox should be in local-space.
            //If e.g. a user changes textbody left and right, changing margin on the parent should not change the Local coordinates
            //Therefore a group in-between should hold the margins
            _marginGroup.Left = LeftMargin;
            _marginGroup.Top = TopMargin;

            var marginTitleItem = new TitleRenderItem("TextBox Margin Group");
            _marginGroup.AddChildItem(marginTitleItem);
        }
        public override void AppendRenderItems(List<Transform> renderItems)
        {
            PrepareForRender();
            renderItems.Add(_group);
            //var rect = Rectangle;

            ////As the rect item is inside the group, we set the left and right to the group and top and left on the rect to 0.
            //_group.Left = Left;
            //_group.Top = Top;
            //_group.Width = Width;
            //_group.Height = Height;

            //_group.TextAnchor = TextAnchor.ToEnumString();
            //renderItems.Add(_group);
            //rect.Top = 0;
            //rect.Left = 0;

            //if (TextBody.AutoSize)
            //{
            //    TextBody.ApplyAutoSize();
            //}

            //rect.Width = Width;
            //rect.Height = Height;

            //var titleItem = new TitleRenderItem("TextBox group");
            //_group.ChildObjects.Add(titleItem);
            ////The rect shound encapse the text element, so we need to set the left depending on the text anchor.
            //if (TextAnchor == eTextAnchor.Middle)
            //{
            //    _group.Left += -(rect.Width / 2);
            //}
            //else if (TextAnchor == eTextAnchor.End)
            //{
            //    if (Math.Abs(Rotation) == 45)
            //    {
            //        const double COS45 = 0.70710678118654757; //Constant for Math.Sin(Math.PI / 4) --45 degrees
            //        _group.Left += -(rect.Width * COS45);
            //        _group.Top += (rect.Width * COS45);
            //    }
            //    else
            //    {
            //        _group.Left += rect.Height / 2;
            //        _group.Top += (rect.Width);
            //    }
            //}
            //_group.ChildObjects.Add(rect);

            ////The textbox should be in local-space.
            ////If e.g. a user changes textbody left and right, changing margin on the parent should not change the Local coordinates
            ////Therefore a group in-between should hold the margins
            //_marginGroup.Left = LeftMargin;
            //_marginGroup.Top = TopMargin;

            //var marginTitleItem = new TitleRenderItem("TextBox Margin Group");
            //_marginGroup.AddChildItem(marginTitleItem);

            //_group.AddChildItem(_marginGroup);
            //TextBody.AppendRenderItems(_marginGroup.ChildObjects);
        }
        /// <summary>
        /// How the text is anchored.
        /// </summary>
        public eTextAnchor TextAnchor
        {
            get;
            set;
        }
    }
}
