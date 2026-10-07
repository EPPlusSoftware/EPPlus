/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  27/11/2025         EPPlus Software AB           EPPlus 9
 *************************************************************************************************/
using EPPlus.Graphics;
using EPPlus.Graphics.Geometry;
using EPPlusImageRenderer;
using System.Drawing;

namespace EPPlus.DrawingRenderer.RenderItems
{
    public enum FillType
    {
        SolidFill,
        GradientFill,
        PatternFill
    }
    /// <summary>
    /// The compound line type. Used for underlining text
    /// </summary>
    public enum CompoundLineStyle
    {        
        /// <summary>
        /// Double lines with equal width
        /// </summary>
        Double,
        /// <summary>
        /// Single line normal width
        /// </summary>
        Single,
        /// <summary>
        /// Double lines, one thick, one thin
        /// </summary>
        DoubleThickThin,
        /// <summary>
        /// Double lines, one thin, one thick
        /// </summary>
        DoubleThinThick,
        /// <summary>
        /// Three lines, thin, thick, thin
        /// </summary>
        TripleThinThickThin
    }
    public enum LineCap
    {
        /// <summary>
        /// A flat line cap
        /// </summary>
        Flat,   //flat
        /// <summary>
        /// A round line cap
        /// </summary>
        Round,
        /// <summary>
        /// A Square line cap
        /// </summary>
        Square
    }

    public enum LineJoin
    {
        Arcs,
        Bevel,
        Miter,
        MiterClip,
        Round
    }
    public class UseReferenceRenderItem : RenderItem
    {
        public UseReferenceRenderItem(BoundingBox parent, string hRef) : base(parent)
        {
            Href = hRef;
        }
        public string Href { get; private set; }

        public override RenderItemType Type => RenderItemType.UseReference;

        public double X
        {
            get
            {
                return Left;
            }
            set
            {
                Left = value;
            }
        }
        public double Y
        {
            get
            {
                return Top;
            }
            set
            {
                Top = value;
            }
        }
        public override RenderItem Clone()
        {
            var clone = new UseReferenceRenderItem((BoundingBox)Parent, Href);
            CloneBase(clone);
            return clone;
        }
    }
    public class RectRenderItem : RenderItem 
    {
        public RectRenderItem() : base()
        {
            
        }
        public RectRenderItem(Transform parent) : base(parent)
        {

        }
        public override RenderItemType Type => RenderItemType.Rect;
        public double RoundedCornerRadius { get; set; }
        public override RenderItem Clone()
        {
            var clone = new RectRenderItem(Parent)
            {
                RoundedCornerRadius = RoundedCornerRadius,
            };

            clone.Width = Width;
            clone.Height = Height;

            CloneBase(clone);
            return clone;
        }
    }
    public class GroupRenderItem : RenderItem
    {
        public GroupRenderItem(BoundingBox parent) : base(parent)
        {
        }
        public GroupRenderItem(BoundingBox parent, double rotation) : base(parent)
        {
            LocalRotation = rotation;
        }

        public GroupRenderItem() : base()
        {
            Parent = TranslationOffset;
        }

        public GroupRenderItem(double localXPos, double localYPos) : this()
        {
            TranslationOffset = new Graphics.TranformPoint(localXPos, localYPos);
        }


        public GroupRenderItem(BoundingBox parent, double rotation, Transform rotationPoint = null) : this(0, 0)
        {
            TranslationOffset.Parent = parent;
            LocalRotation = rotation;
            if (rotationPoint != null)
            {
                RotationPoint = new Graphics.TranformPoint(rotationPoint.LocalPosition.X, rotationPoint.LocalPosition.Y);
            }
        }
        public override RenderItemType Type => RenderItemType.Group;

        //Note: This does not take negative child items into acount
        //TODO: Fix that

        public void AddChildItem(Transform item)
        {
            //item.Bounds.Parent = TranslationOffset;  //This incorrectly sets the parent bounds to zero. Intended?
            ChildObjects.Add(item);

            if (item is BoundingBox bb)
            {
                Width = bb.Right > Width ? bb.Right : Width;
                Height = bb.Bottom > Height ? bb.Bottom : Height;
            }
        }

        ///// <summary>
        ///// Create a subGroup beneath the ParentGroup
        ///// (Or beneath altOverrideBounds but add the renderitems to the parentGroup) This is strange and due to legacy
        ///// </summary>
        ///// <typeparam name="T">Some RenderItem type</typeparam>
        ///// <param name="subGroupName">Class name of the subgroup for easier debugging</param>
        ///// <param name="Items">The RenderItems to place within the group</param>
        ///// <param name="parentGroup">The parent group of this item</param>
        //public void AddSubGroupingOfRenderItems<T>(string subGroupName, List<T> Items) where T : RenderItem
        //{
        //    if (Items != null)
        //    {
        //        //Create subGroup
        //        var subGroup = new GroupRenderItem(this);
        //        subGroup.Name = subGroupName;

        //        //Add items to subGroup
        //        foreach (var renderItem in Items)
        //        {
        //            subGroup.ChildObjects.Add(renderItem);
        //        }

        //        //Add subGroup to parent group
        //        this.ChildObjects.Add(subGroup);
        //    }
        //}


        public override RenderItem Clone()
        {
            var item = new GroupRenderItem(this)
            {
                Rotation = Rotation,
                TextAnchor = TextAnchor,
                TransformOrigin = TransformOrigin,
               // GroupTransform = GroupTransform,
                RotationPoint = RotationPoint,
               // Scale = Scale,
            };
            CloneBase(item);
            foreach(var child in ChildObjects)            
            { 
                item.ChildObjects.Add(((RenderItem)child).Clone());
            }
            return item;
        }


    }
    public class PathRenderItem : RenderItem
    {
        public override RenderItemType Type => RenderItemType.Path;
        public PathRenderItem(Transform parent) : base(parent)
        {

        }
        public List<PathCommand> Commands { get; } = new List<PathCommand>();
        public override RenderItem Clone()
        {
            var clone = new PathRenderItem(Parent);
            CloneBase(clone);
            foreach(var cmd in Commands)
            {
                clone.Commands.Add(cmd.Clone());
            }
            return clone;
        }
    }
    public class EllipseRenderItem : RenderItem
    {
        public EllipseRenderItem(Transform parent) : base(parent)
        {

        }
        public override RenderItemType Type => RenderItemType.Ellipse;
        public double Cx { get; set; }
        public double Cy { get; set; }
        public double Rx { get; set; }
        public double Ry { get; set; }
        public override RenderItem Clone()
        {
            var clone = new EllipseRenderItem(Parent)
            {
                Cx = Cx,
                Cy = Cy,
                Rx = Rx,
                Ry = Ry
            };
            CloneBase(clone);
            return clone;
        }
    }
    public class LineRenderItem : RenderItem
    {
        public LineRenderItem(Transform parent) : base(parent)
        {
            
        }
        double _x1, _y1, _x2, _y2;
        public double X1
        {
            get
            {
                return _x1;
            }
            set
            {
                _x1 = value;
                UpdateBounds();
            }
        }
        public double Y1
        {
            get
            {
                return _y1;
            }
            set
            {
                _y1 = value;
                UpdateBounds();
            }
        }
        public double X2
        {
            get
            {
                return _x2;
            }
            set
            {
                _x2 = value;
                UpdateBounds();
            }
        }
        public double Y2
        {
            get
            {
                return _y2;
            }
            set
            {
                _y2 = value;
                UpdateBounds();
            }
        }
        private void UpdateBounds()
        {
            var px = Math.Min(X1, X2);
            var py = Math.Min(Y1, Y2);
            var sizeX = Math.Abs(X2 - X1);
            var sizeY = Math.Abs(Y2 - Y1);

            LocalPosition = new Vector2(px, py);
            Size = new Vector2(sizeX, sizeY);
        }
        public override RenderItem Clone()
        {
            var clone = new LineRenderItem(Parent);
            CloneBase(clone);
            clone._x1 = X1;
            clone._y1 = Y1;
            clone._x2 = X2;
            clone._y2 = Y2;
            clone.UpdateBounds();
            return clone;
        }
        public override RenderItemType Type => RenderItemType.Line;
    }
    public abstract class DrawingObject
    {
        public virtual void AppendRenderItems(List<Transform> renderItems) { }
    }
    public class RenderItemStyle
    {
        public string DefId { get; set; }
        public string FillColor { get; set; }
        public string FilterName { get; set; }
        public RenderGradientFill GradientFill { get; set; }
        public FillType FillType { get; set; }
        public double? FillOpacity { get; set; }
        public string BorderColor { get; set; }
        public RenderGradientFill BorderGradientFill { get; set; }
        public RenderPatternFill PatternFill { get; set; }
        public RenderBlipFill BlipFill { get; set; }
        public double? BorderWidth { get; set; }
        public double[] BorderDashArray { get; set; }
        public int? StrokeMiterLimit { get; set; }
        public CompoundLineStyle CompoundLineStyle { get; set; } = CompoundLineStyle.Single;
        public double? BorderDashOffset { get; set; }
        public LineCap LineCap { get; set; } = LineCap.Flat;
        public LineJoin LineJoin { get; set; } = LineJoin.Miter;
        public double? BorderOpacity { get; set; }
        public PathFillMode FillColorSource { get; set; } = PathFillMode.Norm;
        public PathFillMode BorderColorSource { get; set; } = PathFillMode.Norm;
        public double? GlowRadius { get; set; }
        public double? GlowOpacity { get; set; }
        public string GlowColor { get; set; }
        public RenderShadowEffect OuterShadowEffect { get; set; }
        internal void GetOuterShadowColor(out string shadowColor, out double opacity)
        {
            if (OuterShadowEffect == null)
            {
                shadowColor = null;
                opacity = 0;

            }
            else
            {
                var tc = OuterShadowEffect.OuterShadowEffectColor;
                if (tc.A < 255 && tc != Color.Empty)
                {
                    opacity = tc.A / 255D;
                }
                else
                {
                    opacity = 1;
                }
                shadowColor = "#" + tc.ToArgb().ToString("x8").Substring(2);
            }
        }

        internal string GetFilterKey()
        {
            return $"{GlowColor} {GlowRadius} {OuterShadowEffect?.GetKey()}";
        }

        internal RenderItemStyle Clone()
        {
            var item = new RenderItemStyle();
            item.FillColor = FillColor;
            item.FillOpacity = FillOpacity;
            item.BorderWidth = BorderWidth;
            item.BorderColor = BorderColor;
            item.BorderDashArray = BorderDashArray;
            item.BorderDashOffset = BorderDashOffset;
            item.BorderOpacity = BorderOpacity;
            item.LineJoin = LineJoin;
            item.LineCap = LineCap;
            item.FillColorSource = FillColorSource;
            return item;
        }
    }
    public abstract class RenderItem : BoundingBox
    {
        protected RenderItem()
        {
        }
        protected RenderItem(Transform parent)
        {
            Parent = parent;
        }
        public abstract RenderItemType Type { get; }
        public RenderItemStyle Style{ get; private set; } = new RenderItemStyle();
        /// <summary>
        /// The origin point for any transform actions in svg.
        /// Normally/Default 0,0
        /// </summary>
        public Coordinate TransformOrigin { get; set; } = null;
        protected void CloneBase(RenderItem item)
        {
            item.Style = Style.Clone();
        }
        public abstract RenderItem Clone();
    }
}