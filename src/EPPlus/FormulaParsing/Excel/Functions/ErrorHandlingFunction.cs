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
using OfficeOpenXml.FormulaParsing.FormulaExpressions;
using OfficeOpenXml.FormulaParsing.Exceptions;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions
{
    /// <summary>
    /// Base class for functions that handles an error that occurs during the
    /// normal execution of the function.
    /// If an exception occurs during the Execute-call that exception will be
    /// caught by the compiler, then the HandleError-method will be called.
    /// </summary>
    public abstract class ErrorHandlingFunction : ExcelFunction
    {
        /// <summary>
        /// Indicates that the function is an ErrorHandlingFunction.
        /// </summary>
        public override bool IsErrorHandlingFunction
        {
            get
            {
                return true;
            }
        }

        /// <summary>
        /// Method that should be implemented to handle the error.
        /// </summary>
        /// <param name="errorCode"></param>
        /// <returns></returns>
        public abstract CompileResult HandleError(string errorCode);
    }
}
