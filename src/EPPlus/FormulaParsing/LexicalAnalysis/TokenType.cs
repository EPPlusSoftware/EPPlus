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
    public enum TokenType
    {
        /// <summary>
        /// The parsed token represents an operator
        /// </summary>
        Operator,
        /// <summary>
        /// The parsed token represents an negator (negates a numeric expression)
        /// </summary>
        Negator,
        /// <summary>
        /// The parsed token represents an opening parenthesis
        /// </summary>
        OpeningParenthesis,
        /// <summary>
        /// The parsed token represents a clising parenthesis
        /// </summary>
        ClosingParenthesis,
        /// <summary>
        /// The parsed token represents a opening enumerable ('{')
        /// </summary>
        OpeningEnumerable,
        /// <summary>
        /// The parsed token represents a closing enumerable ('}')
        /// </summary>
        ClosingEnumerable,
        /// <summary>
        /// The parsed token represents an opening bracket ('[')
        /// </summary>
        OpeningBracket,
        /// <summary>
        /// The parsed token represents a closing bracket (']')
        /// </summary>
        ClosingBracket,
        /// <summary>
        /// The parsed token represents an enumerable
        /// </summary>
        Enumerable,
        /// <summary>
        /// The parsed token represents a comma
        /// </summary>
        Comma,
        /// <summary>
        /// The parsed token represents a semicolon
        /// </summary>
        SemiColon,
        /// <summary>
        /// The parsed token represents a string
        /// </summary>
        String,
        /// <summary>
        /// The parsed token represents content within a string
        /// </summary>
        StringContent,
        /// <summary>
        /// The parsed token represents a worksheet name
        /// </summary>
        WorksheetName,
        /// <summary>
        /// The parsed token represents the content of a worksheet name
        /// </summary>
        WorksheetNameContent,
        /// <summary>
        /// The parsed token represents an integer value
        /// </summary>
        Integer,
        /// <summary>
        /// The parsed token represents a boolean value
        /// </summary>
        Boolean,
        /// <summary>
        /// The parsed token represents a decimal value
        /// </summary>
        Decimal,
        /// <summary>
        /// The parsed token represents a percentage value
        /// </summary>
        Percent,
        /// <summary>
        /// The parsed token represents an excel function
        /// </summary>
        Function,
        /// <summary>
        /// The parsed token represents an excel address
        /// </summary>
        ExcelAddress,
        /// <summary>
        /// The parsed token represents a NameValue
        /// </summary>
        NameValue,
        /// <summary>
        /// The parsed token represents an InvalidReference error (#REF)
        /// </summary>
        InvalidReference,
        /// <summary>
        /// The parsed token represents a Numeric error (#NUM)
        /// </summary>
        NumericError,
        /// <summary>
        /// The parsed tokens represents an Value error (#VAL)
        /// </summary>
        ValueDataTypeError,
        /// <summary>
        /// The parsed token represents the NULL value
        /// </summary>
        Null,
        /// <summary>
        /// The parsed token represent an unrecognized value
        /// </summary>
        Unrecognized,
        /// <summary>
        /// The parsed token represents an R1C1 address
        /// </summary>
        ExcelAddressR1C1
    }
}
