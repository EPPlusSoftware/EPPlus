/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  01/27/2020         EPPlus Software AB       Initial release EPPlus 5
 *************************************************************************************************/
namespace OfficeOpenXml.Drawing
{
    /// <summary>
    /// Specifies the direction(s) in which to flip the gradient while tiling
    /// </summary>
    public enum eTileFlipMode
    {
        /// <summary>
        /// Tiles are not flipped
        /// </summary>
        None,
        /// <summary>
        /// Tiles are flipped horizontally.
        /// </summary>
        X,
        /// <summary>
        /// Tiles are flipped horizontally and Vertically
        /// </summary>
        XY,
        /// <summary>
        /// Tiles are flipped vertically.
        /// </summary>
        Y
    }
}