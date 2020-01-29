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
namespace OfficeOpenXml.Drawing.Style.Coloring
{
    /// <summary>
    /// Datatypes for color transformation types
    /// </summary>
    public enum eColorTransformDataType
    {
        /// <summary>
        /// Percentage
        /// </summary>
        Percentage,
        /// <summary>
        /// Positive percentage
        /// </summary>
        PositivePercentage,
        /// <summary>
        /// Fixed percentage
        /// </summary>
        FixedPercentage,
        /// <summary>
        /// Fixed positive percentage
        /// </summary>
        FixedPositivePercentage,
        /// <summary>
        /// An angel 
        /// </summary>
        Angle,
        /// <summary>
        /// Fixed angle, ranges from -90 to 90   
        /// </summary>
        FixedAngle90,
        /// <summary>
        /// A booleans
        /// </summary>
        Boolean
    }
}