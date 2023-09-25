/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  03/10/2023         EPPlus Software AB       Initial release EPPlus 7
 *************************************************************************************************/
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions
{
    /// <summary>
    /// Options for parsing function arguments to a list of doubles
    /// </summary>
    public class DoubleEnumerableParseOptions
    {
        /// <summary>
        /// Ignore errors in cells
        /// </summary>
        public bool IgnoreErrors
        {
            get; set;
        }

        /// <summary>
        /// Ignore hidden cells
        /// </summary>
        public bool IgnoreHiddenCells
        {
            get; set;
        }

        /// <summary>
        /// Ignore results from underlying SUBTOTAL or AGGREGATE functions
        /// </summary>
        public bool IgnoreNestedSubtotalAggregate
        {
            get; set;
        }

        /// <summary>
        /// Ignore cells with non-numeric values
        /// </summary>
        public bool IgnoreNonNumeric
        {
            get; set;
        }
    }
}
