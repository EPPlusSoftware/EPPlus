using EPPlus.DrawingRenderer.RenderItems.Textbox;
using EPPlus.DrawingRenderer.Svg;
using EPPlus.Fonts.OpenType.Integration.DataHolders;
using EPPlus.Graphics;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlus.DrawingRenderer.RenderItems.SvgPresets
{
    public enum RenderPresets
    {
        Rectangle,
        Circle,
        RotatedContainer,
        TextContainer,
    }

    public class SvgCanvas : GroupRenderItem
    {
        RectRenderItem backgroundRect;

        public SvgCanvas(double width, double height) 
        {
            Width = width;
            Height = height;

            backgroundRect = new RectRenderItem(this);

            backgroundRect.Width = width;
            backgroundRect.Height = height;
            backgroundRect.Style.FillColor = "aliceBlue";
        }

        public string RenderSvgString()
        {
            StringBuilder sb = new StringBuilder();
            var renderOptions = new SvgRenderOptions();
            //renderOptions.Width = (int)Math.Round(bounds.Width.PointToPixel());
            //renderOptions.Height = (int)Math.Round(bounds.Height.PointToPixel());

            var svgShapeRenderer = new SvgShapeRenderer(this, sb, renderOptions);

            svgShapeRenderer.Render(this.ChildObjects);

            var svg = sb.ToString();
            return svg;
        }

        bool rotatedContainerIsAdded = false;
        bool rectAdded = false;
        bool circleAdded = false;

        private void AddRect()
        {
            if(rectAdded)
            {
                return;
            }
            else
            {
                rectAdded = true;
            }

            RectRenderItem rectItem = new RectRenderItem(this);
            rectItem.Style.FillColor = "darkGreen";
            rectItem.Width = 25;
            rectItem.Height = 25;
        }

        private void AddCircle()
        {
            if (circleAdded)
            {
                return;
            }
            else
            {
                circleAdded = true;
            }

            EllipseRenderItem ellipse = new EllipseRenderItem(this);
            ellipse.Style.FillColor = "darkBlue";
            ellipse.Rx = 10;
            ellipse.Ry = 10;
        }

        private void AddRotatingContainer()
        {
            if (rotatedContainerIsAdded)
            {
                return;
            }

            var baseBB = new BoundingBox();

            //96x96 px
            baseBB.Width = 72;
            baseBB.Height = 72;

            BoundingBox parent = new BoundingBox();

            var groupItem = new GroupRenderItem(this, 45);

            groupItem.Left = 10;
            groupItem.Top = 10;

            RectRenderItem rectItem = new RectRenderItem(groupItem);

            rectItem.Style.FillColor = "red";
            rectItem.Style.FillOpacity = 0.2d;

            rectItem.Width = 20;
            rectItem.Height = 20;

            groupItem.AddChildItem(rectItem);


            RectRenderItem siblingItem = new RectRenderItem(groupItem);
            siblingItem.Style.FillColor = "blue";
            siblingItem.Style.FillOpacity = 0.2d;

            siblingItem.Width = 20;
            siblingItem.Height = 20;

            siblingItem.Left = 20;
            siblingItem.Top = 20;

            groupItem.AddChildItem(siblingItem);

            groupItem.SetRotationPointToCenterOfGroup();

            RectRenderItem centerOfGroupMarker = new RectRenderItem(groupItem);
            centerOfGroupMarker.Style.FillColor = "green";
            centerOfGroupMarker.Style.FillOpacity = 0.8d;

            centerOfGroupMarker.Width = 6;
            centerOfGroupMarker.Height = 6;

            centerOfGroupMarker.Left = 30 - (centerOfGroupMarker.Width / 2);
            centerOfGroupMarker.Top = 30 - (centerOfGroupMarker.Height / 2);

            var sb = new StringBuilder();
        }

        private void AddTextContainer(double rotation)
        {
            TextContainer container = new TextContainer(this);

            container.AddText("My Text");
            var rtItem = new RichTextFormatSimple("Second paragraph", "Archivo Narrow", 16f, true);
            rtItem.FontColor = Color.DarkGreen;
            container.AddRichText(rtItem);

            container.Left += Width/2d;
            container.Top += Height/2d;

            container.InnerRotation = rotation;
        }

        public void AddRenderItem(RenderPresets preset, double rotation = 0)
        {
            switch (preset)
            {
                case RenderPresets.Rectangle:
                    AddRect();
                    break;
                case RenderPresets.Circle:
                    AddCircle();
                    break;
                case RenderPresets.RotatedContainer:
                    AddRotatingContainer();
                    break;
                case RenderPresets.TextContainer:
                    AddTextContainer(rotation);
                    break;
            }
        }

        public void ClearChildren()
        {
            backgroundRect.Parent = null;
            base.ChildObjects.Clear();
            backgroundRect.Parent = this;
        }
    }
}
