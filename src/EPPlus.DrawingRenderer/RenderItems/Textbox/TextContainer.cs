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
    public class TextContainer : GroupRenderItem
    {
        RenderContext? _renderContext = null;

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

        double LeftMargin 
        { 
            get { return marginGroup.Left; } 
            set{ marginGroup.Left = value;} 
        }

        double TopMargin
        {
            get { return marginGroup.Top; }
            set { marginGroup.Top = value; }
        }

        public double RightMargin
        {
            get; set;
        }

        public double BottomMargin
        {
            get; set;
        }

        public double Rotation
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
        RectRenderItem topLeftBounds = new RectRenderItem();


        private void init()
        {
            Name = "TextContainer";

            marginGroup = new GroupRenderItem(this);
            marginGroupFill = new RectRenderItem(marginGroup);
            marginGroupFill.Style.FillColor = "blue";
            marginGroupFill.Style.FillOpacity = 0.2d;

            ptOnCircle.Width = 5d;
            ptOnCircle.Height = 5d;
            ptOnCircle.Style.FillColor = "green";
            ptOnCircle.Parent = Parent;

            topLeftBounds.Width = 5d;
            topLeftBounds.Height = 5d;
            topLeftBounds.Style.FillColor = "red";
            topLeftBounds.Parent = Parent;



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

        //double GetXPosOnCircle(double y)
        //{
        //    var arcX = Math.Sqrt(Math.Pow(marginGroupFill.Width, 2) - Math.Pow((y - Position.Y), 2)) + Position.X;
        //    return arcX;
        //}

        //double GetYPosOnCircle(double x)
        //{
        //    var arcY = Math.Sqrt(Math.Pow(marginGroupFill.Width, 2) - Math.Pow((x - Position.X), 2)) + Position.Y;
        //    return arcY;
        //}

        void CalculateMargin()
        {
            marginGroup.Width = txtBody.Width + RightMargin;
            marginGroup.Height = txtBody.Height + BottomMargin;

            marginGroupFill.Width = marginGroup.Width + LeftMargin;
            marginGroupFill.Height = marginGroup.Height + RightMargin;

            //0 when rotation point is top left
            //marginGroupFill.Top = 0;
            marginGroupFill.Left = -marginGroup.Left * -Math.Cos(marginGroupFill.RotationRadians);

            bgFillRect.Width = Width;
            bgFillRect.Height = Height;

            //0 when rotation point is top left
            //marginGroupFill.Top = 0;
            //bgFillRect.Left = -(Math.Cos(marginGroup.LocalRotationRadians) * (marginGroup.Height));
            ptOnCircle.Position = new Graphics.Geometry.Vector2(marginGroupFill.Width * Math.Cos(marginGroupFill.RotationRadians) + Position.X, marginGroupFill.Width * Math.Sin(marginGroupFill.RotationRadians)+ Position.Y);
            topLeftBounds.Position = new Graphics.Geometry.Vector2(marginGroupFill.GetGlobalBoundingbox().Left, marginGroupFill.GetGlobalBoundingbox().Top);
            bgFillRect.Position = topLeftBounds.Position;
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

    }
}
