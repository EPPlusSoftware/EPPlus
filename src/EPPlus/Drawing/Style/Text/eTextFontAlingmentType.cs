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
    /// How characters are aligned vertically.
    /// </summary>
    public enum eTextFontAlingmentType
    {
        /// <summary>
        /// When the text flow is horizontal or simple vertical same as fontBaseline but for other vertical modes same as fontCenter.
        /// </summary>
        Automatic,
        /// <summary>
        ///  The letters are anchored to the very bottom of a single line. This is different than the bottom baseline because of letters such as "g," "q," "y," etc.
        /// </summary>
        Bottom,
        /// <summary>
        ///  The letters are anchored to the bottom baseline of a single line.
        /// </summary>
        Baseline,
        /// <summary>
        ///  The letters are anchored between the two baselines of a single line.
        /// </summary>
        Center,
        /// <summary>
        /// The letters are anchored to the top baseline of a single line.
        /// </summary>
        Top
    }
}