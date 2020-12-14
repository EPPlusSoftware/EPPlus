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
namespace OfficeOpenXml.Drawing.Controls
{
    /// <summary>
    /// A style for a form control drop-down.
    /// </summary>
    public enum eDropStyle
    {
        /// <summary>
        /// A standard combo box
        /// </summary>
        Combo,
        /// <summary>
        /// An editable combo box
        /// </summary>
        ComboEdit,
        /// <summary>
        /// A standard combo box with only the drop-down button visible when the box is not expanded
        /// </summary>
        Simple
    }
}