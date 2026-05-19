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
using System.Drawing;

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
    public class  OffsetRectangle
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

    }
    public class RenderLinearGradientSettings
    {
        public double Angle { get; set; }  
        public bool Scaled { get; set; }
    }
    public class RenderGradientFill
    {
        public RenderGradientFill()
        {
            
        }
        //public RenderGradientFill(List<Color> colors, List<double> stops)
        //{
        //    for (int i = 0; i < stops.Count; i++)
        //    {
        //        var c = new GradientFillColor(stops[i], colors[i]);
        //        Colors.Add(c);
        //    }
        //}

        //public ExcelDrawingGradientFill Settings { get; set; }
        public List<GradientFillColor> Colors { get; set; } = new List<GradientFillColor>();
        public ShadePath ShadePath { get; set; }
        public OffsetRectangle FocusPoint { get; set; } = new OffsetRectangle();
        public OffsetRectangle TileRectangle { get; set; } = new OffsetRectangle();
        public RenderLinearGradientSettings LinearSettings { get; private set; }

    }
}