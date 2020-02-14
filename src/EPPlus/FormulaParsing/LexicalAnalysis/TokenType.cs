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

namespace OfficeOpenXml.FormulaParsing.LexicalAnalysis
{
    /// <summary>
    /// Token types in the context of formula parsing.
    /// </summary>
    [Flags]
    public enum TokenType : uint
    {
        /// <summary>
        /// The parsed token represents an operator
        /// </summary>
        Operator = 1,
        /// <summary>
        /// The parsed token represents an negator (negates a numeric expression)
        /// </summary>
        Negator = 2,
        /// <summary>
        /// The parsed token represents an opening parenthesis
        /// </summary>
        OpeningParenthesis = 4,
        /// <summary>
        /// The parsed token represents a clising parenthesis
        /// </summary>
        ClosingParenthesis = 8,
        /// <summary>
        /// The parsed token represents a opening enumerable ('{')
        /// </summary>
        OpeningEnumerable = 16,
        /// <summary>
        /// The parsed token represents a closing enumerable ('}')
        /// </summary>
        ClosingEnumerable = 32,
        /// <summary>
        /// The parsed token represents an opening bracket ('[')
        /// </summary>
        OpeningBracket = 64,
        /// <summary>
        /// The parsed token represents a closing bracket (']')
        /// </summary>
        ClosingBracket = 128,
        /// <summary>
        /// The parsed token represents an enumerable
        /// </summary>
        Enumerable = 256,
        /// <summary>
        /// The parsed token represents a comma
        /// </summary>
        Comma = 512,
        /// <summary>
        /// The parsed token represents a semicolon
        /// </summary>
        SemiColon = 1024,
        /// <summary>
        /// The parsed token represents a string
        /// </summary>
        String = 2048,
        /// <summary>
        /// The parsed token represents content within a string
        /// </summary>
        StringContent = 4096,
        /// <summary>
        /// The parsed token represents a worksheet name
        /// </summary>
        WorksheetName = 8192,
        /// <summary>
        /// The parsed token represents the content of a worksheet name
        /// </summary>
        WorksheetNameContent = 16384,
        /// <summary>
        /// The parsed token represents an integer value
        /// </summary>
        Integer = 32768,
        /// <summary>
        /// The parsed token represents a boolean value
        /// </summary>
        Boolean = 65536,
        /// <summary>
        /// The parsed token represents a decimal value
        /// </summary>
        Decimal = 131072,
        /// <summary>
        /// The parsed token represents a percentage value
        /// </summary>
        Percent = 262144,
        /// <summary>
        /// The parsed token represents an excel function
        /// </summary>
        Function = 524288,
        /// <summary>
        /// The parsed token represents an excel address
        /// </summary>
        ExcelAddress = 1048576,
        /// <summary>
        /// The parsed token represents a NameValue
        /// </summary>
        NameValue = 2097152,
        /// <summary>
        /// The parsed token represents an InvalidReference error (#REF)
        /// </summary>
        InvalidReference = 4194304,
        /// <summary>
        /// The parsed token represents a Numeric error (#NUM)
        /// </summary>
        NumericError = 8388608,
        /// <summary>
        /// The parsed tokens represents an Value error (#VAL)
        /// </summary>
        ValueDataTypeError = 16777216,
        /// <summary>
        /// The parsed token represents the NULL value
        /// </summary>
        Null = 33554432,
        /// <summary>
        /// The parsed token represent an unrecognized value
        /// </summary>
        Unrecognized = 67108864,
        /// <summary>
        /// The parsed token represents an R1C1 address
        /// </summary>
        ExcelAddressR1C1 = 134217728,
        /// <summary>
        /// The parsed token represents a circular reference
        /// </summary>
        CircularReference = 268435456
    }
}
