using EPPlus.Fonts.OpenType.Utils;
using EPPlus.Graphics;
using EPPlusImageRenderer;
using EPPlusImageRenderer.RenderItems;
using EPPlusImageRenderer.Svg;
using OfficeOpenXml.FormulaParsing.Excel.Functions.MathFunctions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlus.Export.ImageRenderer.RenderItems.Independent.Shared
{
    internal abstract class IndependentRect : RenderItemIndependent
    {
        internal IndependentRect(BoundingBox parent)
        {
            Bounds.Parent = parent;
        }

        internal IndependentRect(BoundingBox parent, BoundingBox bounds) : this(parent)
        {
            Bounds = bounds;
        }

        public double Left { get { return Bounds.Left; } set { Bounds.Left = value; } }
        public double Top { get { return Bounds.Top; } set { Bounds.Top = value; } }
        public double Width { get { return Bounds.Width; } set { Bounds.Width = value; } }
        public double Height { get { return Bounds.Height; } set { Bounds.Height = value; } }
        public double Right { get { return Bounds.Left + Width; } }
        public double Bottom { get { return Bounds.Top + Height; } }
        public double GlobalLeft => Bounds.GlobalLeft;
        public double GlobalTop => Bounds.GlobalTop;
        public double GlobalRight => Bounds.GlobalLeft + Width;
        public double GlobalBottom => Bounds.GlobalTop + Height;

        public override RenderItemType Type => RenderItemType.Rect;


        //public override void Render(StringBuilder sb)
        //{
        //    //var groupItem = new SvgGroupItem(DrawingRenderer, Bounds);
        //    //groupItem.Render(sb);

        //    RenderRect(sb);

        //    //groupItem.RenderEndGroup(sb);
        //}

        //internal void RenderRect(StringBuilder sb)
        //{
        //    sb.AppendFormat("<rect x=\"{0}\" y=\"{1}\" width=\"{2}\" height=\"{3}\" ",
        //        Left.PointToPixelString(),
        //        Top.PointToPixelString(),
        //        Width.PointToPixelString(),
        //        Height.PointToPixelString());
        //    base.Render(sb);
        //    sb.AppendFormat("/>");
        //}

    }
}
