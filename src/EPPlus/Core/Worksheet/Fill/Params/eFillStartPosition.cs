/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
    08/11/2021         EPPlus Software AB       EPPlus 5.8
 *************************************************************************************************/
namespace OfficeOpenXml
{
    /// <summary>
    /// If the fill starts from the top-left cell or the bottom-right cell.
    /// Also see <seealso cref="eFillDirection"/>
    /// </summary>
    public enum eFillStartPosition
    {
        /// <summary>
        /// The fill starts from the top-left cell and fills to the left and down depending on the <see cref="eFillDirection"/>
        /// </summary>
        TopLeft,
        /// <summary>
        /// The fill starts from the bottom-right cell and fills to the right and up depending on the <see cref="eFillDirection"/>
        /// </summary>
        BottomRight
    }
}