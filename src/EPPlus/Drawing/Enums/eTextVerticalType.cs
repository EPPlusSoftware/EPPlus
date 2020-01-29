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
    /// Vertical text type
    /// </summary>
    public enum eTextVerticalType
    {
        /// <summary>
        /// East Asian version of vertical text. Normal fonts are displayed as if rotated by 90 degrees while some East Asian are displayed vertical.
        /// </summary>
        EastAsianVertical,
        /// <summary>
        /// Horizontal text. Default
        /// </summary>
        Horizontal,
        /// <summary>
        /// East asian version of vertical text. . Normal fonts are displayed as if rotated by 90 degrees while some East Asian are displayed vertical. LEFT RIGHT
        /// </summary>
        MongolianVertical,
        /// <summary>
        /// All of the text is vertical orientation, 90 degrees rotated clockwise
        /// </summary>
        Vertical,
        /// <summary>
        /// All of the text is vertical orientation, 90 degrees rotated counterclockwise
        /// </summary>
        Vertical270,
        /// <summary>
        /// All of the text is vertical
        /// </summary>
        WordArtVertical,
        /// <summary>
        /// Vertical WordArt will be shown from right to left
        /// </summary>
        WordArtVerticalRightToLeft

    }
}