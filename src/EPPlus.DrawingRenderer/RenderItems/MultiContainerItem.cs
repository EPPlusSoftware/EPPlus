using EPPlus.Graphics;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlus.DrawingRenderer.RenderItems
{
    public class MultiContainerItem : RenderItem
    {
        public override RenderItemType Type => RenderItemType.Group;

        private Rect ContentBounds {  get; set; }

        public MultiContainerItem(BoundingBox parent) : base(parent)
        {
            ContentBounds = new Rect(0, 0);
        }

        private static void GetExtremeGlobalBounds(Transform t, ref double minLeft, ref double minTop, ref double maxRight, ref double maxBottom)
        {
            if (t is BoundingBox bb)
            {
                if (bb.GlobalRight > maxRight)
                {
                    maxRight = bb.GlobalRight;
                }
                if (bb.GlobalBottom > maxBottom)
                {
                    maxBottom = bb.GlobalBottom;
                }
            }

            if (t.Position.X < minLeft)
            {
                minLeft = t.Position.X;

                if (minLeft > maxRight)
                {
                    maxRight = minLeft;
                }
            }
            if (t.Position.Y < minTop)
            {
                minTop = t.Position.Y;

                if (minTop > maxBottom)
                {
                    maxBottom = minTop;
                }
            }
        }

        private void GetExtremeGlobalBoundsOfAllChildrenRecursive(Transform t, ref double minLeft, ref double minTop, ref double maxRight, ref double maxBottom)
        {
            GetExtremeGlobalBounds(t, ref minLeft, ref minTop, ref maxRight, ref maxBottom);
            if (t.ChildObjects != null && t.ChildObjects.Count > 0)
            {
                foreach (var child in t.ChildObjects)
                {
                    GetExtremeGlobalBoundsOfAllChildrenRecursive(child, ref minLeft, ref minTop, ref maxRight, ref maxBottom);
                }
            }
        }

        List<Transform> cachedChildObjects = null;

        private void UpdateContentBounds()
        {
            if(cachedChildObjects != null)
            {
                if (cachedChildObjects.SequenceEqual(ChildObjects))
                {
                    //No need to update if it's exactly the same
                    //TODO: Verify this is not considered sequence equal if children of childobjects change.
                    return;
                }
            }

            if(ChildObjects != null && ChildObjects.Count > 0)
            {
                var minLeft = double.MaxValue;
                var minTop = double.MaxValue;
                var maxRight = double.MinValue;
                var maxBottom = double.MinValue;

                foreach (var child in ChildObjects) 
                {
                    GetExtremeGlobalBoundsOfAllChildrenRecursive(child, ref minLeft, ref minTop, ref maxRight, ref maxBottom);
                }

                var TopLeftLocal = TransformPointToLocal(new Graphics.Geometry.Vector2(minLeft, minTop));
                var BottomRightLocal = TransformPointToLocal(new Graphics.Geometry.Vector2(maxRight, maxBottom));

                ContentBounds.Left = TopLeftLocal.X;
                ContentBounds.Top = TopLeftLocal.Y;
                ContentBounds.Right = BottomRightLocal.X;
                ContentBounds.Bottom = BottomRightLocal.Y;

                cachedChildObjects = ChildObjects;
            }
        }

        public double ContentTop { get { UpdateContentBounds(); return ContentBounds.Top; } }
        public double ContentLeft { get { UpdateContentBounds(); return ContentBounds.Left; } }
        public double ContentWidth { get { UpdateContentBounds(); return ContentBounds.Width; } }
        public double ContentHeight { get { UpdateContentBounds(); return ContentBounds.Height; } }
        public double ContentRight { get { UpdateContentBounds(); return ContentBounds.Right; } }
        public double ContentBottom { get { UpdateContentBounds(); return ContentBounds.Bottom; } }

        public override RenderItem Clone()
        {
            throw new NotImplementedException();
        }
    }
}
