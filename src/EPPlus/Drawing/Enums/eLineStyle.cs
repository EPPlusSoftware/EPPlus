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
    /// Preset line dash
    /// </summary>
    public enum eLineStyle
    {
        /// <summary>
        /// Dash 1111000
        /// </summary>
        Dash,
        /// <summary>
        /// Dash Dot
        /// 11110001000
        /// </summary>
        DashDot,
        /// <summary>
        /// Dot 1000
        /// </summary>
        Dot,
        /// <summary>
        /// Large Dash 
        ///11111111000
        /// </summary>
        LongDash,
        /// <summary>
        ///  Large Dash Dot 
        ///  111111110001000
        /// </summary>
        LongDashDot,
        /// <summary>
        /// Large Dash Dot Dot
        /// 1111111100010001000
        /// </summary>
        LongDashDotDot,
        /// <summary>
        /// Solid 
        /// 1
        /// </summary>
        Solid,
        /// <summary>
        /// System Dash
        /// 1110
        /// </summary>
        SystemDash,
        /// <summary>
        /// System Dash Dot
        /// 111010
        /// </summary>
        SystemDashDot,
        /// <summary>
        /// System Dash Dot Dot
        /// 11101010
        /// </summary>
        SystemDashDotDot,
        /// <summary>
        /// System Dot 
        /// 10
        /// </summary>
        SystemDot
    }
}