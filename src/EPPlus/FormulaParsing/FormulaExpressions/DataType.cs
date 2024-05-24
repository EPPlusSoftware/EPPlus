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

namespace OfficeOpenXml.FormulaParsing.FormulaExpressions
{
    /// <summary>
    /// Represents a value's data type in the formula parser.
    /// </summary>
    public enum DataType
    {
        /// <summary>
        /// An integer 
        /// </summary>
        Integer = 0,
        /// <summary>
        /// A decimal or floating point
        /// </summary>
        Decimal = 1,
        /// <summary>
        /// A string 
        /// </summary>
        String = 2,
        /// <summary>
        /// A boolean
        /// </summary>
        Boolean = 3,
        /// <summary>
        /// A date or date/time
        /// </summary>
        Date = 4,
        /// <summary>
        /// A time
        /// </summary>
        Time = 5,
        /////// <summary>
        /////// A range or a collection
        /////// </summary>
        ////Enumerable = 6,
        ///// <summary>
        ///// A lookup array
        ///// </summary>
        ////LookupArray = 7,
        /////// <summary>
        /////// A range reference
        /////// </summary>
        ////ExcelAddress = 8,
        /////// <summary>
        /////// Single cell address, e.g A1
        /////// </summary>
        ////ExcelCellAddress = 9,
        /// <summary>
        /// An address range, e.g A1:B2
        /// </summary>
        ExcelRange = 10,
        /// <summary>
        /// An error code
        /// </summary>
        ExcelError = 11,
        /// <summary>
        /// Null or empty string
        /// </summary>
        Empty = 12,
        /// <summary>
        /// An unknown data type
        /// </summary>
        Unknown = 13,
        /// <summary>
        /// Variable data type
        /// </summary>
        Variable = 14
    }
}
