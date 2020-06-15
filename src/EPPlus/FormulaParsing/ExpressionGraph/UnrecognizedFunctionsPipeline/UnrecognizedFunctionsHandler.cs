/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  06/15/2020         EPPlus Software AB       EPPlus 5.2
 *************************************************************************************************/
using OfficeOpenXml.FormulaParsing.Excel.Functions;
using System;
using System.Collections.Generic;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.ExpressionGraph.UnrecognizedFunctionsPipeline
{
    /// <summary>
    /// Examines an unrecognized function name, returns a function if it can be handled
    /// </summary>
    internal abstract class UnrecognizedFunctionsHandler
    {
        /// <summary>
        /// Examines an unrecognized function name, returns a function if it can be handled
        /// </summary>
        /// <param name="funcName">The unrecognized function name</param>
        /// <param name="function">An <see cref="ExcelFunction"/> that can execute the function</param>
        /// <returns></returns>
        public abstract bool Handle(string funcName, out ExcelFunction function);
    }
}
