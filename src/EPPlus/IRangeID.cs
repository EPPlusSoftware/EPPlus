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
using System.Text;

namespace OfficeOpenXml
{
    /// <summary>
    /// Id from a cell, column or row.
    /// </summary>
    interface IRangeID
    {
        /// <summary>
        /// This is the id for a cell, row or column.
        /// The id is a composit of the SheetID, the row number and the column number.
        /// Bit 1-14 SheetID, Bit 15-28 Column number (0 if entire column), Bit 29- Row number (0 if entire row).
        /// </summary>
        ulong RangeID
        {
            get;
            set;
        }
    }
}
