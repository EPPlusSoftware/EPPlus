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
using OfficeOpenXml.FormulaParsing.FormulaExpressions.FunctionCompilers;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions
{
    /// <summary>
    /// Function module
    /// </summary>
    public interface IFunctionModule
    {
        /// <summary>
        /// Gets a dictionary of custom function implementations.
        /// </summary>
        IDictionary<string, ExcelFunction> Functions { get; }

        ///// <summary>
        ///// Gets a dictionary of custom function compilers. A function compiler is not 
        ///// necessary for a custom function, unless the default expression evaluation is not
        ///// sufficient for the implementation of the custom function. When a FunctionCompiler instance
        ///// is created, it should be given a reference to the same function instance that exists
        ///// in the Functions collection of this module.
        ///// </summary>
        //IDictionary<Type, FunctionCompiler> CustomCompilers { get; }
  }
}
