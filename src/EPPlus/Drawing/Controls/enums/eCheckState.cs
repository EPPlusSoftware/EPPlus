/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  10/21/2020         EPPlus Software AB           Controls 
 *************************************************************************************************/
namespace OfficeOpenXml
{
    /// <summary>
    /// The state of a check box form control
    /// </summary>
    public enum eCheckState
    {
        /// <summary>
        /// The checkbox is unchecked
        /// </summary>
        Unchecked,
        /// <summary>
        /// The checkbox is checked
        /// </summary>
        Checked,
        /// <summary>
        /// The checkbox is greyed out, neither checked or unchecked
        /// </summary>
        Mixed
    }
}