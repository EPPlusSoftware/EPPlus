using EPPlus.Graphics;
using EPPlusImageRenderer;
using EPPlusImageRenderer.RenderItems;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Text;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlus.Export.ImageRenderer.RenderItems.Shared
{
    internal abstract class ContainerItem : RenderItem
    {
        private RenderItem InnerItem;
        private RenderItem OuterItem;

        double _marginLeft;
        double _marginTop;

        internal double MarginLeft
        {
            get
            {
                return _marginLeft;
            }
            set
            {
                VerifyMarginInput(value);
                _marginLeft = value;
                //InnerItem.Bounds.Left = value;
            }
        }

        internal double MarginTop
        {
            get
            {
                return _marginTop;
            }
            set
            {
                VerifyMarginInput(value);
                _marginTop = value;
                //InnerItem.Bounds.Top = value;
            }
        }

        double _marginRight;

        internal double MarginRight
        {
            get
            {
                return _marginRight;
            }
            set
            {
                VerifyMarginInput(value);
                _marginRight = value;
                //InnerItem.Bounds.Width = Bounds.Width - value;
            }
        }

        double _marginBottom;

        internal double MarginBottom
        {
            get
            {
                return _marginBottom;
            }
            set
            {
                VerifyMarginInput(value);
                _marginBottom = value;
            }
        }

        bool SizeToContents = true;
        public ContainerItem(RenderItem innerItem, RenderItem outerItem) : base(innerItem.DrawingRenderer)
        {
            OuterItem = outerItem;
            InnerItem = innerItem;

            OuterItem.Bounds.Parent = Bounds;
            InnerItem.Bounds.Parent = OuterItem.Bounds;
        }

        internal void ApplyMargins()
        {
            //Origin-Point: Set. Set in stone. All text moves from there
            //Here it is Bounds.Top and Bounds.Left. Nothing in the content can change that

            OuterItem.Bounds.Top = MarginLeft;
            OuterItem.Bounds.Left = MarginTop;

            OuterItem.Bounds.Width = InnerItem.Bounds.Width + MarginRight;
            OuterItem.Bounds.Height = InnerItem.Bounds.Height + MarginBottom;

            Bounds.Width = OuterItem.Bounds.Width;
            Bounds.Height = OuterItem.Bounds.Height;
        }

        public override void Render(StringBuilder sb)
        {
            OuterItem.Render(sb);
            InnerItem.Render(sb);
        }

        public override RenderItemType Type => RenderItemType.Group;

        internal double GetInnerLeft()
        {
            return Bounds.Left + MarginLeft;
        }

        internal double GetInnerTop()
        {
            return Bounds.Top + MarginTop;
        }

        internal double GetInnerBottom()
        {
            return Bounds.Bottom - MarginBottom;
        }

        internal double GetInnerRight()
        {
            return Bounds.Right - MarginRight;
        }

        public double GetInnerWidth()
        {
            return GetInnerRight() - GetInnerLeft();
        }

        public double GetInnerHeight()
        {
            return GetInnerBottom() - GetInnerTop();
        }

        private void VerifyMarginInput(double value)
        {
            if (value < 0)
            {
                throw new ArgumentException("Margins cannot be set to a negative value! If you wish to set starting position, please set Left or Top for the ContainerItem");
            }
        }
    }
}
