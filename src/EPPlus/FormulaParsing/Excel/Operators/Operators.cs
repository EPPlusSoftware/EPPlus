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

namespace OfficeOpenXml.FormulaParsing.Excel.Operators
{
    /// <summary>
    /// Operator enum
    /// </summary>
    internal enum Operators
    {
        /// <summary>
        /// Undefined
        /// </summary>
        Undefined,
        /// <summary>
        /// Concat
        /// </summary>
        Concat,
        /// <summary>
        /// Plus
        /// </summary>
        Plus,
        /// <summary>
        /// Minus
        /// </summary>
        Minus,
        /// <summary>
        /// Multiply
        /// </summary>
        Multiply,
        /// <summary>
        /// Divide
        /// </summary>
        Divide,
        /// <summary>
        /// Modulus
        /// </summary>
        Modulus,
        /// <summary>
        /// Percent
        /// </summary>
        Percent,
        /// <summary>
        /// Equals
        /// </summary>
        Equals,
        /// <summary>
        /// Greater than
        /// </summary>
        GreaterThan,
        /// <summary>
        /// Greater than or equal
        /// </summary>
        GreaterThanOrEqual,
        /// <summary>
        /// Less than
        /// </summary>
        LessThan,
        /// <summary>
        /// Less than or equal
        /// </summary>
        LessThanOrEqual,
        /// <summary>
        /// Not equal to
        /// </summary>
        NotEqualTo,
        /// <summary>
        /// Integer division
        /// </summary>
        IntegerDivision,
        /// <summary>
        /// Exponentiation
        /// </summary>
        Exponentiation,
        /// <summary>
        /// Colon
        /// </summary>
        Colon,
        /// <summary>
        /// Intersect
        /// </summary>
        Intersect
    }
    /// <summary>
    /// Limited operators
    /// </summary>
    public enum LimitedOperators
    {
        /// <summary>
        /// Equals
        /// </summary>
        Equals = Operators.Equals,
        /// <summary>
        /// Greater than
        /// </summary>
        GreaterThan = Operators.GreaterThan,
        /// <summary>
        /// Greater than or equal
        /// </summary>
        GreaterThanOrEqual = Operators.GreaterThanOrEqual,
        /// <summary>
        /// Less than
        /// </summary>
        LessThan = Operators.LessThan,
        /// <summary>
        /// Less than or equal
        /// </summary>
        LessThanOrEqual = Operators.LessThanOrEqual,
        /// <summary>
        /// Not equal to
        /// </summary>
        NotEqualTo = Operators.NotEqualTo,
    }
}
