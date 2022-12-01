/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  11/07/2022         EPPlus Software AB       Initial release EPPlus 6.2
 *************************************************************************************************/
using OfficeOpenXml.FormulaParsing.ExpressionGraph.FunctionCompilers;
using OfficeOpenXml.FormulaParsing.ExpressionGraph.UnrecognizedFunctionsPipeline;
using System.Linq;
using System;
using System.Collections;
using System.Collections.Generic;
using OfficeOpenXml.FormulaParsing.Exceptions;
using OfficeOpenXml.FormulaParsing.ExpressionGraph.Rpn.FunctionCompilers;

namespace OfficeOpenXml.FormulaParsing.ExpressionGraph.Rpn
{
    internal class RpnFunctionExpression : RpnExpression
    {
        private readonly string _functionName;
        private readonly IEnumerable<RpnExpression> _arguments;
        private readonly RpnFunctionCompilerFactory _functionCompilerFactory;
        private readonly int _pos;
        bool _negate=false;
        internal RpnFunctionExpression(string tokenValue, ParsingContext ctx, int pos) : base(ctx)
        {
            _functionName = tokenValue;
            if (_functionName.StartsWith("_xlfn.", StringComparison.OrdinalIgnoreCase)) _functionName = _functionName.Replace("_xlfn.", string.Empty);
            _arguments = new List<RpnExpression>();
            _pos = pos;
            _functionCompilerFactory = new RpnFunctionCompilerFactory(ctx.Configuration.FunctionRepository, ctx);

        }
        internal override ExpressionType ExpressionType => ExpressionType.Function;
        public override void Negate()
        {
            _negate = !_negate;
        }
        public override CompileResult Compile()
        {
            try
            {
                // older versions of Excel (pre 2007) adds "_xlfn." in front of some function names for compatibility reasons.
                // EPPlus implements most of these functions, so we just remove this.

                var function = Context.Configuration.FunctionRepository.GetFunction(_functionName);
                //if (function == null)
                //{
                //    // Handle unrecognized func name
                //    var pipeline = new FunctionsPipeline(Context, Children);
                //    function = pipeline.FindFunction(_functionName);
                //    if (function == null)
                //    {
                //        if (Context.Debug)
                //        {
                //            Context.Configuration.Logger.Log(Context, string.Format("'{0}' is not a supported function", _functionName));
                //        }
                //        return new CompileResult(ExcelErrorValue.Create(eErrorType.Name), DataType.ExcelError);
                //    }
                //}
                if (Context.Debug)
                {
                    Context.Configuration.Logger.LogFunction(_functionName);
                }
                var compiler = _functionCompilerFactory.Create(function);
                var result = compiler.Compile(_arguments.Any() ? _arguments : Enumerable.Empty<RpnExpression>());
                //if (_isNegated)
                //{
                //    if (!result.IsNumeric)
                //    {
                //        if (_parsingContext.Debug)
                //        {
                //            var msg = string.Format("Trying to negate a non-numeric value ({0}) in function '{1}'",
                //                result.Result, funcName);
                //            _parsingContext.Configuration.Logger.Log(_parsingContext, msg);
                //        }
                //        return new CompileResult(ExcelErrorValue.Create(eErrorType.Value), DataType.ExcelError);
                //    }
                //    return new CompileResult(result.ResultNumeric * -1, result.DataType);
                //}
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
        private RpnExpressionStatus _status= RpnExpressionStatus.NoSet;
        internal override RpnExpressionStatus Status
        {
            get
            {
                if(_status==RpnExpressionStatus.NoSet)
                {
                    foreach(var a in _arguments)
                    {
                        if(a.Status==RpnExpressionStatus.IsAddress)
                        {
                            _status= a.Status;
                            return _status;
                        }
                    }
                    _status= RpnExpressionStatus.CanCompile;
                }
                return _status;
            }
            set
            {
                _status = value;
            }
        }

    }

}
