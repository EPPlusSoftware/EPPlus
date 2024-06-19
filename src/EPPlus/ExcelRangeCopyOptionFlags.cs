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
using System;

namespace OfficeOpenXml
{
    /// <summary>
    /// Flag enum, specify all flags that you want to exclude from the copy.
    /// </summary>
    [Flags]    
    public enum ExcelRangeCopyOptionFlags : int
    {
        /// <summary>
        /// Exclude formulas from being copied. Only the value of the cell will be copied
        /// </summary>
        ExcludeFormulas = 0x1,
        /// <summary>
        /// Will exclude formulas and values from being copied
        /// </summary>
        ExcludeValues = 0x2,
        /// <summary>
        /// Exclude styles from being copied. 
        /// </summary>
        ExcludeStyles = 0x4,
        /// <summary>
        /// Exclude comments from being copied. 
        /// </summary>
        ExcludeComments = 0x8,
        /// <summary>
        /// Exclude threaded comments from being copied. 
        /// </summary>
        ExcludeThreadedComments = 0x10,
        /// <summary>
        /// Exclude hyperlinks from being copied. 
        /// </summary>
        ExcludeHyperLinks = 0x20,
        /// <summary>
        /// Exclude merged cells from being copied. 
        /// </summary>
        ExcludeMergedCells = 0x40,
        /// <summary>
        /// Exclude data validations from being copied. 
        /// </summary>
        ExcludeDataValidations = 0x80,
        /// <summary>
        /// Exclude conditional formatting from being copied. 
        /// </summary>
        ExcludeConditionalFormatting = 0x100,
        /// <summary>
        /// Transpose the copied data
        /// </summary>
        Transpose = 0x200,
        /// <summary>
        /// Exclude drawings from being copied
        /// </summary>
        ExcludeDrawings = 0x400,
    }
}
