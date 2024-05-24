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
using OfficeOpenXml.FormulaParsing.Excel.Functions;
using System.Collections;
using OfficeOpenXml.FormulaParsing.Utilities;

namespace OfficeOpenXml.FormulaParsing.FormulaExpressions.FunctionCompilers
{
    /// <summary>
    /// Function compiler
    /// </summary>
    public abstract class FunctionCompiler
    {
        /// <summary>
        /// Function
        /// </summary>
        protected ExcelFunction Function
        {
            get;
            private set;
        }
        /// <summary>
        /// Function compiler
        /// </summary>
        /// <param name="function">The function</param>
        public FunctionCompiler(ExcelFunction function)
        {
            Require.That(function).Named("function").IsNotNull();
            Function = function;
        }
        /// <summary>
        /// Build function arguments
        /// </summary>
        /// <param name="compileResult"></param>
        /// <param name="dataType"></param>
        /// <param name="args"></param>
        protected void BuildFunctionArguments(CompileResult compileResult, DataType dataType, List<FunctionArgument> args)
        {
            var funcArg = new FunctionArgument(compileResult);
            args.Add(funcArg);
        }
        /// <summary>
        /// Build Function Arguments
        /// </summary>
        /// <param name="result"></param>
        /// <param name="args"></param>
        protected void BuildFunctionArguments(CompileResult result, List<FunctionArgument> args)
        {
            BuildFunctionArguments(result, result.DataType, args);
        }
        /// <summary>
        /// Compile
        /// </summary>
        /// <param name="children"></param>
        /// <param name="context"></param>
        /// <returns></returns>
        public abstract CompileResult Compile(IEnumerable<CompileResult> children, ParsingContext context);

    }
}
