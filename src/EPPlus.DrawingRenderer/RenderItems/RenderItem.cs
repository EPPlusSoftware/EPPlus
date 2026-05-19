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
using EPPlus.Fonts.OpenType.Utils;
using EPPlus.Graphics;
using EPPlus.Graphics.Geometry;
using EPPlusImageRenderer;
using EPPlusImageRenderer.RenderItems;
using System.Drawing;
using System.Globalization;
using System.Text;
using static EPPlus.DrawingRenderer.RenderItems.LineRenderItem;
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
    public class RectRenderItem : RenderItem 
    {
        public RectRenderItem(BoundingBox parent) : base(parent)
        {

        }
        public override RenderItemType Type => RenderItemType.Rect;
        public double Left { get { return Bounds.Left; } set { Bounds.Left = value; } }
        public double Top { get { return Bounds.Top; } set { Bounds.Top = value; } }
        public double Width { get { return Bounds.Width; } set { Bounds.Width = value; } }
        public double Height { get { return Bounds.Height; } set { Bounds.Height = value; } }
        public double Right { get { return Bounds.Left + Width; } }
        public double Bottom { get { return Bounds.Top + Height; } }
        //public double GlobalLeft => Bounds.GlobalLeft;
        //public double GlobalTop => Bounds.GlobalTop;
        //public double GlobalRight => Bounds.GlobalLeft + Width;
        //public double GlobalBottom => Bounds.GlobalTop + Height;
    }
    public class GroupRenderItem : RenderItem
    {
        public GroupRenderItem(BoundingBox parent) : base(parent)
        {
        }
        public GroupRenderItem(BoundingBox parent, double rotation) : base(parent)
        {
            Rotation = rotation;
        }
        public override RenderItemType Type => RenderItemType.Group;
        public string TextAnchor { get; set; }
        public double Rotation { get; set; }
        public string GroupTransform = "";
        public List<RenderItem> Children { get; } = new List<RenderItem>();
    }
    public class PathRenderItem : RenderItem
    {
        public override RenderItemType Type => RenderItemType.Path;
        public PathRenderItem(BoundingBox parent) : base(parent)
        {

        }
        public List<PathCommands> Commands { get; } = new List<PathCommands>();
    }
    public class EllipseRenderItem : RenderItem
    {
        public EllipseRenderItem(BoundingBox parent) : base(parent)
        {

        }
        public override RenderItemType Type => RenderItemType.Rect;
        public double Cx { get; set; }
        public double Cy { get; set; }
        public double Rx { get; set; }
        public double Ry { get; set; }

    }
    public class LineRenderItem : RenderItem
    {
        public LineRenderItem(BoundingBox parent) : base(parent)
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

            Bounds.Position = new Vector2(px, py);
            Bounds.Size = new Vector2(sizeX, sizeY);
        }

        public override RenderItemType Type => RenderItemType.Line;
    }
    public abstract class RenderItem : RenderItemBase
    {
        //internal protected EPPlus.DrawingRenderer DrawingRenderer { get; }
        //internal RenderItem(DrawingB renderer)
        //{
        //    DrawingRenderer = renderer;
        //}

        protected RenderItem(/*DrawingBase renderer,*/ BoundingBox parent)
        {
            Bounds.Parent = parent;
            //DrawingRenderer = renderer; 
        }
        //internal abstract void GetBounds(out double il, out double it, out double ir, out double ib);
        public virtual void GetBounds(out double il, out double it, out double ir, out double ib)
        {
            il = Bounds.Left;
            it = Bounds.Top;
            ir = Bounds.Right;
            ib = Bounds.Bottom;
        }
        internal string DefId = null;
        //internal bool IsEndOfGroup { get; set; } = false;
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
        public int StrokeMiterLimit { get; set; } = 4;
        public CompoundLineStyle CompoundLineStyle { get; set; } = CompoundLineStyle.Single;
        public double? BorderDashOffset { get; set; }
        public LineCap LineCap { get; set; } = LineCap.Flat;
        public LineJoin LineJoin { get; set; } = LineJoin.Miter;
        public double? BorderOpacity { get; set; }
        public PathFillMode FillColorSource { get; set; } = PathFillMode.Norm;
        public PathFillMode BorderColorSource { get; set; } = PathFillMode.Norm;
        public double? GlowRadius { get; set; }
        public string GlowColor { get; set; }
        public RenderShadowEffect OuterShadowEffect { get; private set; } = null;

        /// <summary>
        /// The origin point for any transform actions in svg.
        /// Normally/Default 0,0
        /// </summary>
        public Coordinate TransformOrigin { get; set; } = null;

        protected void CloneBase(RenderItem item)
        {
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
        }
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
    }
    /// <summary>
    /// Base class for any item rendered.
    /// </summary>
    public abstract class RenderItemBase
    {
        public BoundingBox Bounds = new BoundingBox();
        public abstract RenderItemType Type { get; }
        //public abstract void Render(StringBuilder sb);
    }
}