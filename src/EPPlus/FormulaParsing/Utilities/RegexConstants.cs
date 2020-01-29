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

namespace OfficeOpenXml.FormulaParsing.Utilities
{
    /// <summary>
    /// Regex constants for formula parsing.
    /// </summary>
    public static class RegexConstants
    {
        /// <summary>
        /// Regex constant matching a single cell address.
        /// </summary>
        public const string SingleCellAddress = @"^(('[^/\\?*\[\]]{1,31}'|[A-Za-z_]{1,31})!)?[A-Z]{1,3}[1-9]{1}[0-9]{0,7}$";
        /// <summary>
        /// Regex constant matching a full Excel address
        /// </summary>
        public const string ExcelAddress = @"^(('[^/\\?*\[\]]{1,31}'|[A-Za-z_]{1,31})!)?[\$]{0,1}([A-Z]|[A-Z]{1,3}[\$]{0,1}[1-9]{1}[0-9]{0,7})(\:({0,1}[A-Z]|[A-Z]{1,3}[\$]{0,1}[1-9]{1}[0-9]{0,7})){0,1}$";
        //public const string ExcelAddress = @"^([\$]{0,1}([A-Z]{1,3}[\$]{0,1}[0-9]{1,7})(\:([\$]{0,1}[A-Z]{1,3}[\$]{0,1}[0-9]{1,7}){0,1})|([\$]{0,1}[A-Z]{1,3}\:[\$]{0,1}[A-Z]{1,3})|([\$]{0,1}[0-9]{1,7}\:[\$]{0,1}[0-9]{1,7}))$";
        /// <summary>
        /// Regex constant matching a boolean expression (true or false)
        /// </summary>
        public const string Boolean = @"^(true|false)$";
        /// <summary>
        /// Regex constant matching a decimal expression
        /// </summary>
        public const string Decimal = @"^[0-9]+\.[0-9]+$";
        /// <summary>
        /// Regex constant matching an integer expression
        /// </summary>
        public const string Integer = @"^[0-9]+$";
    }
}
