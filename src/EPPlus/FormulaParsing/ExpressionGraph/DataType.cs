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

namespace OfficeOpenXml.FormulaParsing.ExpressionGraph
{
    /// <summary>
    /// Represents a value's data type in the formula parser.
    /// </summary>
    public enum DataType
    {
        /// <summary>
        /// An integer 
        /// </summary>
        Integer,
        /// <summary>
        /// A decimal or floating point
        /// </summary>
        Decimal,
        /// <summary>
        /// A string 
        /// </summary>
        String,
        /// <summary>
        /// A boolean
        /// </summary>
        Boolean,
        /// <summary>
        /// A date or date/time
        /// </summary>
        Date,
        /// <summary>
        /// A time
        /// </summary>
        Time,
        /// <summary>
        /// A range or a collection
        /// </summary>
        Enumerable,
        /// <summary>
        /// A lookup array
        /// </summary>
        LookupArray,
        /// <summary>
        /// A range reference
        /// </summary>
        ExcelAddress,
        /// <summary>
        /// Single cell address, e.g A1
        /// </summary>
        ExcelCellAddress,
        /// <summary>
        /// An address range, e.g A1:B2
        /// </summary>
        ExcelRange,
        /// <summary>
        /// An error code
        /// </summary>
        ExcelError,
        /// <summary>
        /// Null or empty string
        /// </summary>
        Empty,
        /// <summary>
        /// An unknown data type
        /// </summary>
        Unknown,
        /// <summary>
        /// Worksheet name
        /// </summary>
        WorksheetName,
    }
}
