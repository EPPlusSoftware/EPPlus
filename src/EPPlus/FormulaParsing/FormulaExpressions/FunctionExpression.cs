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
using System.Linq;
using System;
using System.Collections;
using System.Collections.Generic;
using OfficeOpenXml.FormulaParsing.Exceptions;
using OfficeOpenXml.FormulaParsing.FormulaExpressions.FunctionCompilers;
using OfficeOpenXml.FormulaParsing.Excel.Functions;

namespace OfficeOpenXml.FormulaParsing.FormulaExpressions
{
    internal enum ExpressionCondition : byte
    {
        None = 0xFF,
        False = 0,
        True = 1,
        Both = 2
    }
    internal class FunctionExpression : Expression
    {
        private FunctionCompilerFactory _functionCompilerFactory;
        internal ExcelFunction _function;
        internal int _startPos, _endPos;
        internal IList<int> _arguments;
        internal int _argPos=0;
        internal ExpressionCondition _latestConitionValue = ExpressionCondition.None;
        int _negate = 0;
        internal FunctionExpression(string tokenValue, ParsingContext ctx, int pos) : base(ctx)
        {

            if (tokenValue.StartsWith("_xlfn.", StringComparison.OrdinalIgnoreCase)) tokenValue = tokenValue.Replace("_xlfn.", string.Empty);
            _arguments = new List<int>();
            _startPos = pos;
            _functionCompilerFactory = new FunctionCompilerFactory(ctx.Configuration.FunctionRepository, ctx);
            _function = ctx.Configuration.FunctionRepository.GetFunction(tokenValue);
            //var compiler = _functionCompilerFactory.Create(_function);
        }
        private FunctionExpression(ParsingContext ctx) : base(ctx)
        {

        }
        internal override ExpressionType ExpressionType => ExpressionType.Function;
        public override void Negate()
        {
            if (_negate == 0)
            {
                _negate = -1;
            }
            else
            {
                _negate *= -1;
            }
        }
        IList<Expression> _args=null;
        internal void SetArguments(IList<Expression> args)
        {
            _args = args;
        }
        public override CompileResult Compile()
        {
            try
            {
                if (Context.Debug)
                {
                    Context.Configuration.Logger.LogFunction(_function.GetType().Name);
                }
                
                if(_function==null) return new CompileResult(ExcelErrorValue.Create(eErrorType.Name), DataType.ExcelError);

                var compiler = _functionCompilerFactory.Create(_function);
                var result = compiler.Compile(_args??Enumerable.Empty<Expression>());
                if (_negate!=0)
                {
                    if (result.IsNumeric==false)
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

        internal int GetTokenPosForArg(FunctionParameterInformation type)
        {
            var i = _argPos;
            while (i < _arguments.Count && _function.GetParameterInfo(i) != type) i++;
            if(i < _arguments.Count ) return _arguments[i];
            return -1;
        }
        internal override Expression CloneWithOffset(int row, int col)
        {
            if(_function==null || _function.HasNormalArguments)
            {
                return this;
            }
            return new FunctionExpression(Context) { _arguments = _arguments, _function = _function, _functionCompilerFactory = _functionCompilerFactory, _startPos = _startPos, _endPos = _endPos  };
        }
        private ExpressionStatus _status= ExpressionStatus.NoSet;
        internal override ExpressionStatus Status
        {
            get
            {
                if(_status==ExpressionStatus.NoSet)
                {
                    //foreach(var a in _arguments)
                    //{
                    //    if(a.Status==RpnExpressionStatus.IsAddress)
                    //    {
                    //        _status= a.Status;
                    //        return _status;
                    //    }
                    //}
                    _status= ExpressionStatus.CanCompile;
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
