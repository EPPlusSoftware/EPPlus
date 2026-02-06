using EPPlus.Fonts.OpenType.Utils;
using EPPlus.Graphics;
using EPPlusImageRenderer;
using EPPlusImageRenderer.RenderItems;
using EPPlusImageRenderer.Svg;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Text;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Xml.Serialization;
using static System.Net.Mime.MediaTypeNames;

namespace EPPlus.Export.ImageRenderer.RenderItems.SvgItem
{
    internal class SvgTextBoxItem : DrawingObject
    {
        internal SvgTextBoxItem(DrawingBase renderer, BoundingBox parent, double left, double top, double width, double height, double maxWidth = double.NaN, double maxHeight = double.NaN) : base(renderer, parent)
        {
            Bounds.Left = left;
            Bounds.Top = top;
            Bounds.Width = width;
            Bounds.Height = height;
            Init(renderer, parent, maxWidth, maxHeight);
        }

        private void Init(DrawingBase renderer, BoundingBox parent, double maxWidth, double maxHeight)
        {
            Bounds.Name = "TextBox";

            TextBody = new SvgTextBodyItem(renderer, Bounds, true);
            TextBody.MaxWidth = maxWidth;
            TextBody.MaxHeight = maxHeight;

            Rectangle = new SvgRenderRectItem(renderer, parent);
            Rectangle.Bounds = Bounds;
        }

        internal SvgTextBoxItem(DrawingBase renderer, BoundingBox parent, double maxWidth, double maxHeight) : base(renderer, parent)
        {
            Init(renderer, parent, maxWidth, maxHeight);
        }

        //Simplified input
        internal SvgTextBoxItem(DrawingBase renderer, BoundingBox parent, BoundingBox maxBounds) : this(
                                            renderer, parent, maxBounds.Left, maxBounds.Top, maxBounds.Width, maxBounds.Height)
        {
        }
        
        public SvgRenderItem Rectangle { get; set; }
        public SvgTextBodyItem TextBody {get;set;}
        double _leftMargin;
        internal double LeftMargin 
        {
            get 
            {
                return _leftMargin;
            }
            set 
            {
                Bounds.Width += (value - _leftMargin);
                _leftMargin = value;
            }
        }

        double _topMargin;
        internal double TopMargin 
        {
            get
            {
                return _topMargin;
            }
            set
            {
                Bounds.Height += (value - _topMargin);
                _topMargin = value;
            }
        }

        double _rightMargin;
        internal double RightMargin
        {
            get
            {
                return _rightMargin;
            }
            set
            {
                Bounds.Width += (value - _rightMargin);
                _rightMargin = value;
            }
        }
        double _bottomMargin;
        internal double BottomMargin
        {
            get
            {
                return _bottomMargin;
            }
            set
            {
                Bounds.Height += (value - _bottomMargin);
                _bottomMargin = value;
            }
        }

        internal void ImportTextBody(ExcelTextBody body, bool autoSize)
        {
            double l, r, t, b;
            body.GetInsetsOrDefaults(out l, out t, out r, out b);
            LeftMargin = l;
            TopMargin = t;
            RightMargin = r;
            BottomMargin = b;

            TextBody.ImportTextBody(body);

            if (autoSize)
            {
                Bounds.Width = TextBody.Bounds.Width + LeftMargin + RightMargin;
                Bounds.Height = TextBody.Bounds.Height + TopMargin + BottomMargin;
            }
        }


        internal override void AppendRenderItems(List<RenderItem> renderItems)
        {
            SvgGroupItem groupItem;
            if (Bounds.Rotation == 0)
            {
                groupItem = new SvgGroupItem(DrawingRenderer, Bounds);
            }
            else
            {
                groupItem = new SvgGroupItem(DrawingRenderer, Bounds, Bounds.Rotation);
            }
            renderItems.Add(groupItem);
            //handled by group now
            Rectangle.Bounds.Top = 0;
            Rectangle.Bounds.Left = 0;
            renderItems.Add(Rectangle);
            TextBody.AppendRenderItems(renderItems);
            renderItems.Add(new SvgEndGroupItem(DrawingRenderer, Bounds));
        }

        internal void ImportParagraph(ExcelDrawingParagraph item, double startingY, string text = null)
        {
            TextBody.ImportParagraph(item, startingY, text);
            Bounds.Width = TextBody.Width;
            Bounds.Height = TextBody.Height;
        }

        internal void ImportTextBody(ExcelTextBody textBody)
        {
            TextBody.ImportTextBody(textBody);
            Bounds.Width = TextBody.Width;
            Bounds.Height = TextBody.Height;
        }
    }
}
