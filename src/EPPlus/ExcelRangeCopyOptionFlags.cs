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
        /// Exclude formulas from being copied
        /// </summary>
        ExcludeFormulas = 0x1,
        ExcludeFormulasAndValues = 0x2,
        ExcludeStyles = 0x4,
        ExcludeComments = 0x8,
        ExcludeThreadedComments = 0x10,
        ExcludeHyperLinks = 0x20,
        ExcludeMergedCells = 0x30,
    }
}
