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

namespace OfficeOpenXml.FormulaParsing.Exceptions
{
    /// <summary>
    /// This Exception represents an Excel error. When this exception is thrown
    /// from an Excel function, the ErrorValue code will be set as the value of the
    /// parsed cell.
    /// </summary>
    /// <seealso cref="ExcelErrorValue"/>
    public class ExcelErrorValueException : Exception
    {
        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="error">The error value causing the exception</param>
        public ExcelErrorValueException(ExcelErrorValue error)
            : this(error.ToString(), error)
        {
            
        }
        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="error">The error value causing the exception</param>
        /// <param name="message">An error message for the exception</param>
        public ExcelErrorValueException(string message, ExcelErrorValue error)
            : base(message)
        {
            ErrorValue = error;
        }
        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="errorType">The error type causing the exception</param>
        public ExcelErrorValueException(eErrorType errorType)
            : this(ExcelErrorValue.Create(errorType))
        {
            
        }

        /// <summary>
        /// The error value
        /// </summary>
        public ExcelErrorValue ErrorValue { get; private set; }
    }
}
