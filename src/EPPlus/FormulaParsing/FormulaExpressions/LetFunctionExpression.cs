using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml.FormulaParsing.Exceptions;

namespace OfficeOpenXml.FormulaParsing.FormulaExpressions
{
    internal class LetFunctionExpression : VariableFunctionExpression
    {
        internal LetFunctionExpression(string tokenValue, ParsingContext ctx, int pos) : base(tokenValue, ctx, pos)
        {

        }


        public override CompileResult Compile()
        {
            try
            {
                if (Context.Debug)
                {
                    Context.Configuration.Logger.LogFunction(_function.GetType().Name);
                }

                if (_function == null) return new CompileResult(ExcelErrorValue.Create(eErrorType.Name), DataType.ExcelError);

                var compiler = Context.FunctionCompilerFactory.Create(_function, Context);
                var result = compiler.Compile(_args ?? Enumerable.Empty<CompileResult>(), Context);

                if (_negate != 0)
                {
                    if (result.IsNumeric == false)
                    {
                        if (Context.Debug)
                        {
                            var msg = string.Format("Trying to negate a non-numeric value ({0}) in function '{1}'",
                                result.Result, nameof(_function));
                            Context.Configuration.Logger.Log(Context, msg);
                        }
                        return new CompileResult(ExcelErrorValue.Create(eErrorType.Value), DataType.ExcelError);
                    }
                    return new CompileResult(result.ResultNumeric * _negate, result.DataType);
                }
                return result;
            }
            catch (ExcelErrorValueException e)
            {
                if (Context.Debug)
                {
                    Context.Configuration.Logger.Log(Context, e);
                }
                return new CompileResult(e.ErrorValue, DataType.ExcelError);
            }
        }
    }
}
