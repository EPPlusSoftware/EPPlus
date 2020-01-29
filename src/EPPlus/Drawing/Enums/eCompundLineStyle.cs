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
    /// The compound line type. Used for underlining text
    /// </summary>
    public enum eCompundLineStyle
    {
        /// <summary>
        /// Double lines with equal width
        /// </summary>
        Double,
        /// <summary>
        /// Single line normal width
        /// </summary>
        Single,
        /// <summary>
        /// Double lines, one thick, one thin
        /// </summary>
        DoubleThickThin,
        /// <summary>
        /// Double lines, one thin, one thick
        /// </summary>
        DoubleThinThick,
        /// <summary>
        /// Three lines, thin, thick, thin
        /// </summary>
        TripleThinThickThin
    }
}