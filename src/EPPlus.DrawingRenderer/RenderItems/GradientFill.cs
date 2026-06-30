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
using EPPlus.Export.ImageRenderer.RenderItems;

using System.Text;

namespace EPPlus.DrawingRenderer.RenderItems
{
    /// <summary>
    /// The path for a gradiant color
    /// </summary>
    public enum ShadePath
    {
        /// <summary>
        /// The gradient folows a linear path
        /// </summary>
        Linear,
        /// <summary>
        /// The gradient follows a circular path
        /// </summary>
        Circle,
        /// <summary>
        /// The gradient follows a rectangular path
        /// </summary>
        Rectangle,
        /// <summary>
        /// The gradient follows the shape
        /// </summary>
        Shape
    }
    public class  OffsetRectangle : RenderStyle
    {
        /// <summary>
        /// Top offset in percentage
        /// </summary>
        public double TopOffset { get; set; }
        /// <summary>
        /// Bottom offset in percentage
        /// </summary>
        public double BottomOffset { get; set; }
        /// <summary>
        /// Left offset in percentage
        /// </summary>
        public double LeftOffset { get; set; }
        /// <summary>
        /// Right offset in percentage
        /// </summary>
        public double RightOffset { get; set; }

        public override string GetKey()
        {
            return $"{TopOffset} {BottomOffset} {LeftOffset} {RightOffset}";
        }
    }
    public class RenderLinearGradientSettings : RenderStyle
    {
        public double Angle { get; set; }  
        public bool Scaled { get; set; }

        public override string GetKey()
        {
            return $"{Angle} {Scaled}";
        }
    }
    public class RenderGradientFill : RenderStyle
    {
        public RenderGradientFill()
        {
            
        }
        public List<GradientFillColor> Colors { get; set; } = new List<GradientFillColor>();
        public ShadePath ShadePath { get; set; }
        public OffsetRectangle FocusPoint { get; set; } = new OffsetRectangle();
        public OffsetRectangle TileRectangle { get; set; } = new OffsetRectangle();
        public RenderLinearGradientSettings LinearSettings { get; private set; } = new RenderLinearGradientSettings();
        /// <summary>
        /// If the gradient should use the user space as coordinate system or the bounding box of the item. This is only used for gradient fills and is ignored for other fill types. If true, the gradient will use the user space as coordinate system, if false, the gradient will use the bounding box of the item as coordinate system.
        /// </summary>
        public bool UserSpaceOnUse { get; set; }
        public override string GetKey()
        {
            var sb = new StringBuilder();
            foreach(var c in Colors)
            {
                sb.Append(c.Color.ToArgb());
                sb.Append(' ');
                sb.Append(c.Position);
            }
            sb.Append(ShadePath);

            sb.Append(' ');

            sb.Append(FocusPoint.GetKey());
            sb.Append(' ');
            sb.Append(LinearSettings.GetKey());

            return sb.ToString();
        }
    }

    public abstract class RenderStyle
    {
        public abstract string GetKey();
    }
}