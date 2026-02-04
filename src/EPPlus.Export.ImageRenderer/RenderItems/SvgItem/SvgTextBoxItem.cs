using EPPlus.Fonts.OpenType.Utils;
using EPPlus.Graphics;
using EPPlusImageRenderer;
using EPPlusImageRenderer.RenderItems;
using EPPlusImageRenderer.Svg;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Text;
using System.Collections.Generic;
using System.Drawing;
using System.Xml.Serialization;

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
            Bounds.Name = "TextBox";

            TextBody = new SvgTextBodyItem(renderer, Bounds, true);
            TextBody.MaxWidth = maxWidth;
            TextBody.MaxHeight = maxHeight;
            //Bounds.Width = TextBody.Width;
            //Bounds.Height = TextBody.Height;

            Rectangle = new SvgRenderRectItem(renderer, parent);
            Rectangle.Bounds = Bounds;
        }
        internal SvgTextBoxItem(DrawingBase renderer, BoundingBox parent, double maxWidth, double maxHeight) : base(renderer, parent)
        {
        }

        //Simplified input
        internal SvgTextBoxItem(DrawingBase renderer, BoundingBox parent, BoundingBox maxBounds) : this(
                                            renderer, parent, maxBounds.Left, maxBounds.Top, maxBounds.Width, maxBounds.Height)
        {
        }

        public SvgRenderItem Rectangle { get; set; }
        public SvgTextBodyItem TextBody {get;set;}
        internal double LeftMargin { get { return TextBody.Bounds.Left; } set { TextBody.Bounds.Left = value; } }

        internal double TopMargin { get { return TextBody.Bounds.Top; } set { TextBody.Bounds.Top = value; } }

        internal double RightMargin { get { return Bounds.Width - TextBody.Bounds.Left - TextBody.Bounds.Width ; } set { TextBody.Bounds.Width = Bounds.Width - TextBody.Bounds.Left - value; } }
        internal double BottomMargin { get { return Bounds.Height - TextBody.Bounds.Top - TextBody.Bounds.Height; } set { TextBody.Bounds.Height = Bounds.Height - TextBody.Bounds.Top - value; } }
        internal void ImportTextBody(ExcelTextBody body)
        {
            double l, r, t, b;
            body.GetInsetsOrDefaults(out l, out t, out r, out b);
            LeftMargin = l.PointToPixel();
            TopMargin = t.PointToPixel();
            RightMargin = r.PointToPixel();
            BottomMargin = b.PointToPixel();

            TextBody.ImportTextBody(body);
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
    }
}
