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
    /// Text alignment
    /// </summary>
    public enum eTextAlignment
    {
        /// <summary>
        /// Left alignment
        /// </summary>
        Left,
        /// <summary>
        /// Center alignment
        /// </summary>
        Center,
        /// <summary>
        /// Right alignment
        /// </summary>
        Right,
        /// <summary>
        /// Distributes the text words across an entire text line
        /// </summary>
        Distributed,
        /// <summary>
        /// Align text so that it is justified across the whole line.
        /// </summary>
        Justified,
        /// <summary>
        /// Aligns the text with an adjusted kashida length for Arabic text
        /// </summary>
        JustifiedLow,
        /// <summary>
        /// Distributes Thai text specially, specially, because each character is treated as a word
        /// </summary>
        ThaiDistributed
    }
}