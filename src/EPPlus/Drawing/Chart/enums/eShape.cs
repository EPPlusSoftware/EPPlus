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
namespace OfficeOpenXml.Drawing.Chart
{
    /// <summary>
    /// Shape for bar charts
    /// </summary>
    public enum eShape
    {
        /// <summary>
        /// A box shape
        /// </summary>
        Box,
        /// <summary>
        /// A cone shape
        /// </summary>
        Cone,
        /// <summary>
        /// A cone shape, truncated to max
        /// </summary>
        ConeToMax,
        /// <summary>
        /// A cylinder shape
        /// </summary>
        Cylinder,
        /// <summary>
        /// A pyramid shape
        /// </summary>
        Pyramid,
        /// <summary>
        /// A pyramid shape, truncated to max
        /// </summary>
        PyramidToMax
    }
}