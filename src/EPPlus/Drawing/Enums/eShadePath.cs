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
    /// The path for a gradiant color
    /// </summary>
    public enum eShadePath
    {
        Linear,
        /// <summary>
        /// Gradient follows a circular path
        /// </summary>
        Circle,
        /// <summary>
        /// Gradient follows a rectangular path
        /// </summary>
        Rectangle, 
        /// <summary>
        /// Gradient follows the shape
        /// </summary>
        Shape 
    }
}