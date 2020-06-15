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
using System.Globalization;
using System.Linq;
using System.Text;
using OfficeOpenXml.FormulaParsing.Excel;
using OfficeOpenXml.FormulaParsing.Excel.Functions;
using OfficeOpenXml.FormulaParsing.Exceptions;
using OfficeOpenXml.FormulaParsing.ExpressionGraph.FunctionCompilers;
using OfficeOpenXml.FormulaParsing.ExpressionGraph.UnrecognizedFunctionsPipeline;
using OfficeOpenXml.Utils;

namespace OfficeOpenXml.FormulaParsing.ExpressionGraph
{
    /// <summary>
    /// Expression that handles execution of a function.
    /// </summary>
    public class FunctionExpression : AtomicExpression
    {
        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="expression">should be the of the function</param>
        /// <param name="parsingContext"></param>
        /// <param name="isNegated">True if the numeric result of the function should be negated.</param>
        public FunctionExpression(string expression, ParsingContext parsingContext, bool isNegated)
            : base(expression)
        {
            _parsingContext = parsingContext;
            _functionCompilerFactory = new FunctionCompilerFactory(parsingContext.Configuration.FunctionRepository, parsingContext);
            _isNegated = isNegated;
            base.AddChild(new FunctionArgumentExpression(this));
        }

        private readonly ParsingContext _parsingContext;
        private readonly FunctionCompilerFactory _functionCompilerFactory;
        private readonly bool _isNegated;


        public override CompileResult Compile()
        {
            try
            {
                var funcName = ExpressionString;

                // older versions of Excel (pre 2007) adds "_xlfn." in front of some function names for compatibility reasons.
                // EPPlus implements most of these functions, so we just remove this.
                if (funcName.StartsWith("_xlfn.")) funcName = funcName.Replace("_xlfn.", string.Empty);

                var function = _parsingContext.Configuration.FunctionRepository.GetFunction(funcName);
                if (function == null)
                {
                    var pipeline = new FunctionsPipeline();
                    function = pipeline.FindFunction(funcName);
                    if(function == null)
                    {
                        if (_parsingContext.Debug)
                        {
                            _parsingContext.Configuration.Logger.Log(_parsingContext, string.Format("'{0}' is not a supported function", funcName));
                        }
                        return new CompileResult(ExcelErrorValue.Create(eErrorType.Name), DataType.ExcelError);
                    }
                }
                if (_parsingContext.Debug)
                {
                    _parsingContext.Configuration.Logger.LogFunction(funcName);
                }
                var compiler = _functionCompilerFactory.Create(function);
                var result = compiler.Compile(HasChildren ? Children : Enumerable.Empty<Expression>());
                if (_isNegated)
                {
                    if (!result.IsNumeric)
                    {
                        if (_parsingContext.Debug)
                        {
                            var msg = string.Format("Trying to negate a non-numeric value ({0}) in function '{1}'",
                                result.Result, funcName);
                            _parsingContext.Configuration.Logger.Log(_parsingContext, msg);
                        }
                        return new CompileResult(ExcelErrorValue.Create(eErrorType.Value), DataType.ExcelError);
                    }
                    return new CompileResult(result.ResultNumeric * -1, result.DataType);
                }
                return result;
            }
            catch (ExcelErrorValueException e)
            {
                if (_parsingContext.Debug)
                {
                    _parsingContext.Configuration.Logger.Log(_parsingContext, e);
                }
                return new CompileResult(e.ErrorValue, DataType.ExcelError);
            }
            
        }

        public override Expression PrepareForNextChild()
        {
            return base.AddChild(new FunctionArgumentExpression(this));
        }

        public override bool HasChildren
        {
            get
            {
                return (Children.Any() && Children.First().Children.Any());
            }
        }

        public override Expression AddChild(Expression child)
        {
            Children.Last().AddChild(child);
            return child;
        }
    }
}
