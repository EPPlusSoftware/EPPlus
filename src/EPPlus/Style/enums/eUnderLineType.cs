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
namespace OfficeOpenXml.Style
{
    /// <summary>
    /// Linestyle
    /// </summary>
    public enum eUnderLineType
    {
        /// <summary>
        /// Dashed
        /// </summary>
        Dash,
        /// <summary>
        /// Dashed, Thicker
        /// </summary>
        DashHeavy,
        /// <summary>
        /// Dashed Long
        /// </summary>
        DashLong,
        /// <summary>
        /// Long Dashed, Thicker
        /// </summary>
        DashLongHeavy,
        /// <summary>
        /// Double lines with normal thickness
        /// </summary>
        Double,
        /// <summary>
        /// Dot Dash
        /// </summary>
        DotDash,
        /// <summary>
        /// Dot Dash, Thicker
        /// </summary>
        DotDashHeavy,
        /// <summary>
        /// Dot Dot Dash
        /// </summary>
        DotDotDash,
        /// <summary>
        /// Dot Dot Dash, Thicker
        /// </summary>
        DotDotDashHeavy,
        /// <summary>
        /// Dotted
        /// </summary>
        Dotted,
        /// <summary>
        /// Dotted, Thicker
        /// </summary>
        DottedHeavy,
        /// <summary>
        /// Single line, Thicker
        /// </summary>
        Heavy,
        /// <summary>
        /// No underline
        /// </summary>
        None,
        /// <summary>
        /// Single line
        /// </summary>
        Single,
        /// <summary>
        /// A single wavy line
        /// </summary>
        Wavy,
        /// <summary>
        /// A double wavy line
        /// </summary>
        WavyDbl,
        /// <summary>
        /// A single wavy line, Thicker
        /// </summary>
        WavyHeavy,
        /// <summary>
        /// Underline just the words and not the spaces between them
        /// </summary>
        Words
    }
}