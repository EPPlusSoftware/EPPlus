/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
    9/11/2025         EPPlus Software AB       EPPlus 9
 *************************************************************************************************/
namespace OfficeOpenXml.Drawing
{
    /// <summary>
    /// Text tab alignment within a single paragraph.
    /// </summary>
    public enum eTabStopParagraphAlignment
    {
        /// <summary>
        /// The text at this tab stop is center aligned.
        /// </summary>
        Center,
        /// <summary>
        /// At this tab stop, the decimals are lined up. From a user's point of view, the text here behaves as right aligned until the decimal, and then as left aligned after the decimal.
        /// </summary>
        Decimal,
        /// <summary>
        /// The text at this tab stop is left aligned.
        /// </summary>
        Left,
        /// <summary>
        /// The text at this tab stop is right aligned.
        /// </summary>
        Right
    }
}