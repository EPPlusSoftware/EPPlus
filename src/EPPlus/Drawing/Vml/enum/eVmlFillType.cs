/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  11/18/2020         EPPlus Software AB       EPPlus 5.5
 *************************************************************************************************/
namespace OfficeOpenXml.Drawing.Vml
{
    /// <summary>
    /// Type of fill style for a vml drawing.
    /// </summary>
    public enum eVmlFillType
    {
        /// <summary>
        /// No fill is applied.
        /// </summary>
        NoFill,
        /// <summary>
        /// The fill pattern is solid.Default
        /// </summary>
        Solid,
        /// <summary>
        /// The fill colors blend together in a linear gradient from bottom to top.
        /// </summary>
        Gradient,
        /// <summary>
        ///  The fill colors blend together in a radial gradient.
        /// </summary>
        GradientRadial,
        /// <summary>
        ///  The fill image is tiled.
        /// </summary>
        Tile,
        /// <summary>
        /// The image is used to create a pattern using the fill colors.
        /// </summary>
        Pattern,
        /// <summary>
        /// The image is stretched to fill the shape.
        /// </summary>
        Frame
    }
}
