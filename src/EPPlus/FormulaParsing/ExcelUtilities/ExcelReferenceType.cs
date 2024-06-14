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
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.ExcelUtilities
{
    /// <summary>
    /// Reference types for if an adress/cell is absolute or relative and in what way
    /// </summary>
    public enum ExcelReferenceType
    {
        /// <summary>
        /// Both Row and column are absolute
        /// </summary>
        AbsoluteRowAndColumn = 1,
        /// <summary>
        /// Absolute row and relative column
        /// </summary>
        AbsoluteRowRelativeColumn = 2,
        /// <summary>
        /// Realtive row absolute column
        /// </summary>
        RelativeRowAbsoluteColumn = 3,
        /// <summary>
        /// Relative row and relative column
        /// </summary>
        RelativeRowAndColumn = 4
    }
}
