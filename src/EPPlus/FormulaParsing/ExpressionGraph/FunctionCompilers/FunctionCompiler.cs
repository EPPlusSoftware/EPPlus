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

namespace OfficeOpenXml.FormulaParsing.ExpressionGraph.FunctionCompilers
{
    public abstract class FunctionCompiler
    {
        protected ExcelFunction Function
        {
            get;
            private set;
        }

        protected ParsingContext Context
        {
            get;
            private set;
        }

        public FunctionCompiler(ExcelFunction function, ParsingContext context)
        {
            Require.That(function).Named("function").IsNotNull();
            Require.That(context).Named("context").IsNotNull();
            Function = function;
            Context = context;
        }

        protected void BuildFunctionArguments(CompileResult compileResult, DataType dataType, List<FunctionArgument> args)
        {
            if (compileResult.Result is IEnumerable<object> && !(compileResult.Result is IRangeInfo))
            {
                var argList = new List<FunctionArgument>();
                var objects = compileResult.Result as IEnumerable<object>;
                foreach (var arg in objects)
                {
                    var cr = CompileResultFactory.Create(arg);
                    BuildFunctionArguments(cr, dataType, argList);
                }
                args.Add(new FunctionArgument(argList));
            }
            else
            {
                var funcArg = new FunctionArgument(compileResult.Result, dataType);
                funcArg.ExcelAddressReferenceId = compileResult.ExcelAddressReferenceId;
                funcArg.Address = compileResult.Address;
                if(compileResult.IsHiddenCell)
                {
                    funcArg.SetExcelStateFlag(Excel.ExcelCellState.HiddenCell);
                }
                args.Add(funcArg);
            }
        }

        protected void BuildFunctionArguments(CompileResult result, List<FunctionArgument> args)
        {
            BuildFunctionArguments(result, result.DataType, args);
        }

        public abstract CompileResult Compile(IEnumerable<Expression> children);

    }
}
