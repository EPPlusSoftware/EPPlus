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
    /// Axis label position
    /// </summary>
    public enum eTickLabelPosition
    {
        /// <summary>
        /// The axis labels will be at the high end of the perpendicular axis
        /// </summary>
        High,
        /// <summary>
        /// The axis labels will be at the low end of the perpendicular axis
        /// </summary>
        Low,
        /// <summary>
        /// The axis labels will be next to the axis.
        /// </summary>
        NextTo,
        /// <summary>
        /// No axis labels are drawn
        /// </summary>
        None
    }
}