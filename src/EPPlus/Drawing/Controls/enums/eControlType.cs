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
namespace OfficeOpenXml.Drawing.Controls
{
    /// <summary>
    /// Type of form control
    /// </summary>
    public enum eControlType
    {
        /// <summary>
        /// A button
        /// </summary>
        Button,
        /// <summary>
        /// A check box
        /// </summary>
        CheckBox,
        /// <summary>
        /// A combo box
        /// </summary>
        DropDown,
        /// <summary>
        /// A group box
        /// </summary>
        GroupBox,
        /// <summary>
        /// A label
        /// </summary>
        Label,
        /// <summary>
        /// A list box
        /// </summary>
        ListBox,
        /// <summary>
        /// An radio button (option button)
        /// </summary>
        RadioButton,
        /// <summary>
        /// A scroll bar
        /// </summary>
        ScrollBar,
        /// <summary>
        /// A spin button
        /// </summary>
        SpinButton,
        /// <summary>
        /// An edit box. Unsupported. Editboxes can't be used directly on a form
        /// </summary>
        EditBox,
        /// <summary>
        /// A dialog. Unsupported.
        /// </summary>
        Dialog
    }

}
