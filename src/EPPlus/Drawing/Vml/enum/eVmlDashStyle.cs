/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  11/23/2020         EPPlus Software AB       EPPlus 5.5
 *************************************************************************************************/
namespace OfficeOpenXml
{
    /// <summary>
    /// Dash style for a line used in VML drawings
    /// </summary>
    public enum eVmlDashStyle
    {
        /// <summary>
        /// A solid line
        /// </summary>
        Solid,
        /// <summary>
        /// Short - Dash
        /// </summary>
        ShortDash,
        /// <summary>
        /// Short - Dot
        /// </summary>
        ShortDot,
        /// <summary>
        /// Short - Dash - Dot
        /// </summary>
        ShortDashDot,
        /// <summary>
        /// Short - Dash - Dot - Dot
        /// </summary>
        ShortDashDotDot,
        /// <summary>
        /// Dotted
        /// </summary>
        Dot,
        /// <summary>
        /// Dashed
        /// </summary>
        Dash,
        /// <summary>
        /// Long dashes
        /// </summary>
        LongDash,
        /// <summary>
        /// Dash - Dot
        /// </summary>
        DashDot,
        /// <summary>
        /// Long Dash - Dot
        /// </summary>
        LongDashDot,
        /// <summary>
        /// Long Dash - Dot - Dot
        /// </summary>
        LongDashDotDot,
        /// <summary>
        /// Custom dash style.
        /// </summary>
        Custom
    }
}