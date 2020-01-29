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
    /// The color type
    /// </summary>
    public enum eDrawingColorType
    {
        /// <summary>
        /// Not specified
        /// </summary>
        None,
        /// <summary>
        /// RGB specified in percentage
        /// </summary>
        RgbPercentage,      //ScRgbColor
        /// <summary>
        /// Red Green Blue
        /// </summary>
        Rgb,
        /// <summary>
        /// Hue, Saturation, Luminance
        /// </summary>
        Hsl,
        /// <summary>
        /// A system color
        /// </summary>
        System,
        /// <summary>
        /// A color bound to a user's theme
        /// </summary>
        Scheme,
        /// <summary>
        /// A preset Color
        /// </summary>
        Preset,
        /// <summary>
        /// A Color refering to a charts color style
        /// </summary>
        ChartStyleColor
    }
}