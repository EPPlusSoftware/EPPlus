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
}