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
    /// Horizontal alignment for a form control. Unused in Excel 2010, so internal for now.
    /// </summary>
    internal enum eHorizontalAlignmentControl
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
        /// Justify alignment
        /// </summary>
        Justify,
        /// <summary>
        /// Distributed alignment
        /// </summary>
        Distributed
    }
}