using EPPlus.DrawingRenderer.RenderItems.SvgItem;
using EPPlus.Fonts.OpenType;
using EPPlus.Graphics;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using EPPlus.Fonts.OpenType.Integration.DataHolders;

namespace EPPlus.DrawingRenderer.RenderItems.Textbox
{
    public enum RotationAnchor
    {
        TopLeft = 1,
        Middle = 2
    }

    public class TextContainer : GroupRenderItem
    {
        RenderContext? _renderContext = null;

        private RotationAnchor _rotationAnchor = RotationAnchor.TopLeft; 

        public TextContainer()
        {
            init();
        }

        public TextContainer(BoundingBox parent) : base(parent)
        {
            init();
        }

        public TextContainer(BoundingBox parent, double rotation) : base(parent, rotation)
        {
            init();
        }

        public TextContainer(double localXPos, double localYPos) : base(localXPos, localYPos)
        {
            init();
        }

        public TextContainer(BoundingBox parent, double rotation, Transform rotationPoint = null) : base(parent, rotation, rotationPoint)
        {
            init();
        }

        RenderTextBody txtBody;
        GroupRenderItem marginGroup;
        RectRenderItem marginGroupFill;

        RectRenderItem bgFillRect;

        Rect GlobalDynamicBounds
        {
            get
            {
                return txtBody.GetGlobalBoundingbox();
            }
        }

        public virtual double LeftMargin
        {
            get { return marginGroup.Left; }
            set { marginGroup.Left = value; CalculateMargin(); }
        }

        public double TopMargin
        {
            get { return marginGroup.Top; }
            set { marginGroup.Top = value; CalculateMargin(); }
        }

        public double RightMargin
        {
            get; set;
        }

        public double BottomMargin
        {
            get; set;
        }

        public virtual double InnerRotation
        {
            get
            {
                return marginGroup.LocalRotation;
            }
            set
            {
                marginGroup.LocalRotation = value;
                CalculateMargin();
            }
        }

        RectRenderItem ptOnCircle = new RectRenderItem();
        RectRenderItem rotPointDot = new RectRenderItem();


        private void init()
        {
            Name = "TextContainer";

            marginGroup = new GroupRenderItem(this);
            marginGroup.Name = "MarginGroup";
            marginGroupFill = new RectRenderItem(marginGroup);
            marginGroupFill.Style.FillColor = "blue";
            marginGroupFill.Style.FillOpacity = 0.2d;
            marginGroupFill.Name = "MarginGroupFill";

            ptOnCircle.Width = 5d;
            ptOnCircle.Height = 5d;
            ptOnCircle.Style.FillColor = "green";
            ptOnCircle.Parent = Parent;

            rotPointDot.Width = 5d;
            rotPointDot.Height = 5d;
            rotPointDot.Style.FillColor = "red";
            rotPointDot.Parent = this;
            //Center object
            rotPointDot.Left -= Width / 2d;
            rotPointDot.Top -= Height / 2d;



            if (_renderContext == null)
            {
                var engine = new OpenTypeFontEngine(x =>
                {
                    x.SearchSystemDirectories = true;
                });
                _renderContext = new RenderContext(() => engine);
            }

            bgFillRect = new RectRenderItem(this);
            bgFillRect.Style.FillColor = "GoldenRod";
            bgFillRect.Style.FillOpacity = 0.2d;

            txtBody = new SvgTextBodyRenderItem(_renderContext, marginGroup, true);
            txtBody.MaxWidth = Parent.Size.X;
            txtBody.MaxHeight = Parent.Size.Y;
        }

        void SetTxtBodyMaxSize(double height, double width)
        {
            txtBody.MaxHeight = height;
            txtBody.MaxWidth = width;
        }

        void CalculateMargin()
        {
            marginGroup.Width = txtBody.Width + RightMargin;
            marginGroup.Height = txtBody.Height + BottomMargin;

            marginGroupFill.Width = marginGroup.Width + LeftMargin;
            marginGroupFill.Height = marginGroup.Height + TopMargin;


            bgFillRect.Width = Width;
            bgFillRect.Height = Height;

            if (_rotationAnchor == RotationAnchor.TopLeft)
            {
                marginGroupFill.Parent = marginGroup;
                //0 when rotation point is top left
                //marginGroupFill.Top = 0;
                marginGroupFill.Left = -marginGroup.Left * Math.Abs(Math.Cos(marginGroupFill.RotationRadians));
                marginGroupFill.Top = -marginGroup.Top * Math.Abs(Math.Cos(marginGroupFill.RotationRadians));

                rotPointDot.Position = marginGroup.Position;
                bgFillRect.Position = new Graphics.Geometry.Vector2(marginGroupFill.GetGlobalBoundingbox().Left, marginGroupFill.GetGlobalBoundingbox().Top);
                rotPointDot.Position = bgFillRect.Position;
            }
            else if (_rotationAnchor == RotationAnchor.Middle)
            {
                marginGroup.SetRotationPointToCenterOfGroup();
                rotPointDot.Position = marginGroup.RotationPoint.Position;

                var rotPointInternal = new TranformPoint(marginGroup.RotationPoint.Left + marginGroup.Position.X, marginGroup.RotationPoint.Top + marginGroup.Position.Y);
                marginGroupFill.Parent = rotPointInternal;

                rotPointDot.Position = rotPointInternal.Position;

                var prevPos = marginGroupFill.Position;

                marginGroupFill.Left = -marginGroupFill.Width / 2d;
                marginGroupFill.Top = -marginGroupFill.Height / 2d;

                rotPointInternal.Rotation = marginGroup.LocalRotation;

                bgFillRect.Position = new Graphics.Geometry.Vector2(marginGroupFill.GetGlobalBoundingbox().Left, marginGroupFill.GetGlobalBoundingbox().Top);


                marginGroupFill.Parent = marginGroup;
                marginGroupFill.Left = 0;
                marginGroupFill.Top = 0;
            }
        }

        public double Width
        {
            get
            {
                //It does not add its own leftmargin to its width
                var prevMargin = LeftMargin;
                marginGroup.Width += prevMargin;
                marginGroup.Left -= prevMargin;
                var rotatedAndScaledWidth = marginGroup.GetGlobalBoundingbox().Width;
                marginGroup.Left = prevMargin;

                return rotatedAndScaledWidth;
            }
        }

        public double Height
        {
            get
            {
                //It does not add its own topMargin to its height
                var prevMargin = TopMargin;
                marginGroup.Height += prevMargin;
                marginGroup.Top -= prevMargin;
                var rotatedAndScaledWidth = marginGroup.GetGlobalBoundingbox().Height;
                marginGroup.Top = prevMargin;

                return rotatedAndScaledWidth;
            }
        }

        public void AddText(string text)
        {
            txtBody.AddParagraph(text);
            txtBody.RecalculateParagraphs();
            CalculateMargin();
        }

        public void AddRichText(IRichTextFormatSimple rt)
        {
            txtBody.AddParagraph(rt);
            txtBody.RecalculateParagraphs();
            CalculateMargin();
        }


        //i = 0 rotation around topLeft, i = 1 rotation midPoint
        public void SetRotation(RotationAnchor anchor)
        {
            switch (anchor)
            {
                case RotationAnchor.TopLeft:
                    _rotationAnchor = anchor;
                    marginGroup.RotationPoint = null;
                    rotPointDot.Position = marginGroup.Position;
                    break;
                case RotationAnchor.Middle:
                    _rotationAnchor = anchor;
                    marginGroup.SetRotationPointToCenterOfGroup();
                    rotPointDot.Position = marginGroup.RotationPoint.Position;
                    break;
            }
        }
    }
}
