using EPPlus.DrawingRenderer.RenderItems.SvgItem;
using EPPlus.Fonts.OpenType;
using EPPlus.Graphics;
using EPPlus.Fonts.OpenType.Integration.DataHolders;
using System.ComponentModel;

namespace EPPlus.DrawingRenderer.RenderItems.Textbox
{
    public enum RotationAnchor
    {
        TopLeft = 1,
        Middle = 2
    }

    public class TextContainer : GroupRenderItem, INotifyPropertyChanged
    {
        RenderContext? _renderContext = null;

        private RotationAnchor _rotationAnchor = RotationAnchor.TopLeft;

        //Currently there MUST be a parent or everything will nullref
        //public TextContainer()
        //{
        //    init();
        //}

        public TextContainer(BoundingBox parent, RenderContext context) : base(parent)
        {
            _renderContext = context;
            init();
        }

        public TextContainer(BoundingBox parent) : base(parent)
        {
            init();
            TextBody = new SvgTextBodyRenderItem(_renderContext, _marginGroup, true);
            TextBody.MaxWidth = Parent.Size.X;
            TextBody.MaxHeight = Parent.Size.Y;
        }

        public TextContainer(BoundingBox parent, double rotation) : base(parent, rotation)
        {
            init();
        }

        //Currently there MUST be a parent or everything will nullref
        //public TextContainer(double localXPos, double localYPos) : base(localXPos, localYPos)
        //{
        //    init();
        //}

        public TextContainer(BoundingBox parent, double rotation, Transform rotationPoint = null) : base(parent, rotation, rotationPoint)
        {
            init();
        }

        /// <summary>
        /// Lock this container at a certain point Y the top of a rotated object will always be placed at this Y no matter the rotation
        /// If == double.NaN no locking/adjustment is done. No locking is default
        /// </summary>
        double TopPositionLock = double.NaN;

        protected RenderTextBody _textBody;

        public virtual RenderTextBody TextBody{ get { return _textBody; } set { _textBody = (SvgTextBodyRenderItem)value; } }
        protected GroupRenderItem _marginGroup;

        protected GroupRenderItem _innerGroup;
        protected GroupRenderItem _rotationGroup;

        RectRenderItem marginGroupFill;

        RectRenderItem bgFillRect;

        public Transform Parent { get { return base.Parent; } set { var hadParent = base.Parent != null; base.Parent = value; if (hadParent) { CalculateMargin(); } } }

        public override double Left { get => base.Left; set { base.Left = value; CalculateMargin(); } }
        public override double Top { get => base.Top; set { base.Top = value; CalculateMargin(); } }

        public virtual double LeftMargin
        {
            get { return _marginGroup.Left; }
            set { _marginGroup.Left = value; CalculateMargin(); }
        }

        public double TopMargin
        {
            get { return _marginGroup.Top; }
            set { _marginGroup.Top = value; CalculateMargin(); }
        }

        double _rightMargin;
        double _bottomMargin;

        public double RightMargin
        {
            get { return _rightMargin; } 
            set { _rightMargin = value; CalculateMargin(); }
        }

        public double BottomMargin
        {
            get { return _bottomMargin; }
            set { _bottomMargin = value; CalculateMargin(); }
        }

        public virtual double InnerRotation
        {
            get
            {
                return _rotationGroup.LocalRotation;
            }
            set
            {
                _rotationGroup.LocalRotation = value;
                CalculateMargin();
            }
        }

        RectRenderItem ptOnCircle = new RectRenderItem();
        RectRenderItem rotPointDot = new RectRenderItem();

        public event PropertyChangedEventHandler? PropertyChanged;

        private void init()
        {
            Name = "TextContainer";


            _innerGroup = new GroupRenderItem(this);
            _innerGroup.Name = "innerGroup/Position adjust";

            _rotationGroup = new GroupRenderItem(_innerGroup);
            _rotationGroup.Name = "Rotation";

            _marginGroup = new GroupRenderItem(_rotationGroup);
            _marginGroup.Name = "MarginGroup";
            marginGroupFill = new RectRenderItem(_marginGroup);
            marginGroupFill.Style.FillColor = "blue";
            //marginGroupFill.Style.FillOpacity = 0.2d;
            marginGroupFill.Style.FillOpacity = 0d;
            marginGroupFill.Name = "MarginGroupFill";
            

            ptOnCircle.Width = 2.5d;
            ptOnCircle.Height = 2.5d;
            ptOnCircle.Style.FillColor = "green";
            ptOnCircle.Style.FillOpacity = 0d;
            ptOnCircle.Parent = this;

            rotPointDot.Width = 2.5d;
            rotPointDot.Height = 2.5d;
            //rotPointDot.Style.FillColor = "red";
            rotPointDot.Style.FillOpacity = 0d;
            rotPointDot.Parent = this;

            rotPointDot.Position = _marginGroup.Position;
            rotPointDot.Left -= LeftMargin;
            rotPointDot.Top -= TopMargin;
            //Center object
            rotPointDot.Left -= rotPointDot.Width / 2d;
            rotPointDot.Top -= rotPointDot.Height / 2d;



            ptOnCircle.Position = new Graphics.Geometry.Vector2(rotPointDot.Position.X, rotPointDot.Position.Y);



            if (_renderContext == null)
            {
                var engine = new OpenTypeFontEngine(x =>
                {
                    x.SearchSystemDirectories = true;
                });
                _renderContext = new RenderContext(() => engine);
            }

            bgFillRect = new RectRenderItem(_innerGroup);
            bgFillRect.Style.FillColor = "GoldenRod";
            bgFillRect.Style.FillOpacity = 0.2d;
        }

        void SetTxtBodyMaxSize(double height, double width)
        {
            TextBody.MaxHeight = height;
            TextBody.MaxWidth = width;
        }

        public double GetActualHeight()
        {
            return marginGroupFill.Height;
        }

        public double GetActualWidth()
        {
            return marginGroupFill.Width;
        }

        protected void CalculateMargin()
        {
            ptOnCircle.Left = _rotationGroup.Position.X;
            ptOnCircle.Top = _rotationGroup.Position.Y;

            _marginGroup.Width = TextBody.Width + RightMargin;
            _marginGroup.Height = TextBody.Height + BottomMargin;

            marginGroupFill.Width = _marginGroup.Width + LeftMargin;
            marginGroupFill.Height = _marginGroup.Height + TopMargin;


            bgFillRect.Width = Width;
            bgFillRect.Height = Height;

            //Fill/bounds should include the margin
            marginGroupFill.Parent = _marginGroup;
            marginGroupFill.Left = -_marginGroup.Left;
            marginGroupFill.Top = -_marginGroup.Top;

            if (_rotationAnchor == RotationAnchor.TopLeft)
            {
                rotPointDot.Left = _marginGroup.GlobalLeft;
                rotPointDot.Top = _marginGroup.GlobalTop;
                //Center object
                rotPointDot.Left -= rotPointDot.Width / 2d;
                rotPointDot.Top -= rotPointDot.Height / 2d;


                var globalBounds = marginGroupFill.GetGlobalBoundingbox();
                bgFillRect.Left = 0;
                bgFillRect.Top = 0;
                bgFillRect.Position = new Graphics.Geometry.Vector2(globalBounds.Left, globalBounds.Top);
                

                if (globalBounds.Width > bgFillRect.Width)
                {
                    bgFillRect.Width = globalBounds.Width;
                }
                if(globalBounds.Height > bgFillRect.Height)
                {
                    bgFillRect.Height = globalBounds.Height;
                }
            }
            else if (_rotationAnchor == RotationAnchor.Middle)
            {
                //Set center of rotation
                marginGroupFill.SetRotationPointToCenterOfGroup();
                var rotationPointForlater = new Graphics.Geometry.Vector2(marginGroupFill.RotationPoint.Left, marginGroupFill.RotationPoint.Top);

                //Use the local rotation point to Calculate global rotation point for internal calculations
                var rotPointInternal = new TranformPoint(marginGroupFill.RotationPoint.Left + marginGroupFill.Position.X, marginGroupFill.RotationPoint.Top + marginGroupFill.Position.Y);
                rotPointDot.Position = rotPointInternal.Position;

                //Rotate internal rotation and parent to it
                rotPointInternal.LocalRotation = _rotationGroup.LocalRotation;
                marginGroupFill.Parent = rotPointInternal;
                //Center the fill
                marginGroupFill.Left = -marginGroupFill.Width / 2d;
                marginGroupFill.Top = - marginGroupFill.Height / 2d;
                marginGroupFill.RotationPoint = null;

                //Calculate bounds/position when rotated
                var globalBounds = marginGroupFill.GetGlobalBoundingbox();
                bgFillRect.Position = new Graphics.Geometry.Vector2(globalBounds.Left, globalBounds.Top);
                if (globalBounds.Width > bgFillRect.Width)
                {
                    bgFillRect.Width = globalBounds.Width;
                }
                if (globalBounds.Height > bgFillRect.Height)
                {
                    bgFillRect.Height = globalBounds.Height;
                }

                //Reset rotations since svg functions differently and we would prefer to rotate from textContainer
                rotPointInternal.LocalRotation = 0;
                marginGroupFill.Parent = _marginGroup;
                marginGroupFill.Left = -_marginGroup.Left;
                marginGroupFill.Top = -_marginGroup.Top;

                //Add rotation point to actual textcontainer instead
                _rotationGroup.RotationPoint = new TranformPoint(rotationPointForlater.X, rotationPointForlater.Y);
            }
        }

        public override double Width
        {
            get
            {
                //It does not add its own leftmargin to its width
                var prevMargin = LeftMargin;
                _marginGroup.Width += prevMargin;
                _marginGroup.Left -= prevMargin;
                var rotatedAndScaledWidth = _marginGroup.GetGlobalBoundingbox().Width;
                _marginGroup.Width -= prevMargin;
                _marginGroup.Left = prevMargin;

                return rotatedAndScaledWidth;
            }
        }

        public override double Height
        {
            get
            {
                //It does not add its own topMargin to its height
                var prevMargin = TopMargin;
                _marginGroup.Height += prevMargin;
                _marginGroup.Top -= prevMargin;
                var rotatedAndScaledWidth = _marginGroup.GetGlobalBoundingbox().Height;
                _marginGroup.Top = prevMargin;
                _marginGroup.Height -= prevMargin;

                return rotatedAndScaledWidth;
            }
        }

        public void AddText(string text)
        {
            TextBody.AddParagraph(text);
            TextBody.RecalculateParagraphs();
            CalculateMargin();
        }

        public void AddRichText(IRichTextFormatSimple rt)
        {
            TextBody.AddParagraph(rt);
            TextBody.RecalculateParagraphs();
            CalculateMargin();
        }


        //i = 0 rotation around topLeft, i = 1 rotation midPoint
        public void SetRotation(RotationAnchor anchor)
        {
            switch (anchor)
            {
                case RotationAnchor.TopLeft:
                    _rotationAnchor = anchor;
                    _marginGroup.RotationPoint = null;
                    rotPointDot.Position = _marginGroup.Position;
                    break;
                case RotationAnchor.Middle:
                    _rotationAnchor = anchor;
                    _marginGroup.SetRotationPointToCenterOfGroup();
                    rotPointDot.Position = _marginGroup.RotationPoint.Position;
                    break;
            }
            CalculateMargin();
        }

        public void AlignTopWithYPos()
        {
            var bounds = marginGroupFill.GetGlobalBoundingbox();
            var positionY = _innerGroup.Position.Y;

            if(positionY > bounds.Top)
            {
                var dist = positionY - bounds.Top;
                _innerGroup.Position = new Graphics.Geometry.Vector2(Position.X, positionY + dist);
            }
            else
            {
                var dist = bounds.Top - positionY;
                _innerGroup.Position = new Graphics.Geometry.Vector2(Position.X, positionY - dist);
            }
        }

        public void AlignBottomWithYPos()
        {
            var bounds = marginGroupFill.GetGlobalBoundingbox();
            var positionY = Position.Y;

            if (positionY > bounds.Bottom)
            {
                var dist = positionY - bounds.Bottom;
                _innerGroup.Position = new Graphics.Geometry.Vector2(_innerGroup.Position.X, _innerGroup.Position.Y + dist);
            }
            else
            {
                var dist = bounds.Bottom - positionY;
                _innerGroup.Position = new Graphics.Geometry.Vector2(_innerGroup.Position.X, _innerGroup.Position.Y - dist);
            }
        }

        public RenderItemStyle GetStyleBg()
        {
            return bgFillRect.Style;
        }
    }
}
