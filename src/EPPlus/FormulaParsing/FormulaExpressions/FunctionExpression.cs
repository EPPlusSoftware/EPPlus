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
using OfficeOpenXml.FormulaParsing.Excel.Functions;
using System.Text;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
using OfficeOpenXml.Utils;
using System.Runtime.Serialization.Formatters;

namespace OfficeOpenXml.FormulaParsing.FormulaExpressions
{
    internal enum ExpressionCondition : byte
    {
        None = 0xFF,
        False = 0,
        True = 1,
        Error = 2,
        Multi = 4
    }
    internal class FunctionExpression : Expression
    {
        internal ExcelFunction _function;
        internal int _startPos, _endPos;
        protected IList<int> _arguments;
        internal int _argPos=0;
        internal ExpressionCondition _latestConditionValue = ExpressionCondition.None;
        internal CompileResult _cachedResult;
        internal int _negate = 0;
        internal FunctionExpression(string tokenValue, ParsingContext ctx, int pos) : base(ctx)
        {

            if (tokenValue.StartsWith("_xlfn.", StringComparison.OrdinalIgnoreCase)) tokenValue = tokenValue.Replace("_xlfn.", string.Empty);
            if (tokenValue.StartsWith("_xlws.", StringComparison.OrdinalIgnoreCase)) tokenValue = tokenValue.Replace("_xlws.", string.Empty);
            _arguments = new List<int>();
            _startPos = pos;
            _function = ctx.Configuration.FunctionRepository.GetFunction(tokenValue);
        }
        private FunctionExpression(ParsingContext ctx) : base(ctx)
        {

        }
        internal override ExpressionType ExpressionType => ExpressionType.Function;

        internal virtual bool HandlesVariables => false;

        internal virtual bool IsVariableArg(int arg, bool isLastArgument)
        {
            return false;
        }

        internal virtual bool IsVariable(string name)
        {
            return false;
        }

        internal virtual void AddArgument(int arg)
        {
            _arguments.Add(arg);
        }

        internal int NumberOfArguments
        {
            get { return _arguments.Count; }
        }

        internal int GetArgument(int arg)
        {
            return _arguments[arg];
        }

        public override Expression Negate()
        {
            if (_negate == 0)
            {
                _negate = -1;
            }
            else
            {
                _negate *= -1;
            }
            return this;
        }
        protected IList<CompileResult> _args=null;
        internal Queue<FormulaRangeAddress> _dependencyAddresses = null;
        internal bool SetArguments(IList<CompileResult> argsResults)
        {
            _args = argsResults;
            if (_function.ParametersInfo.HasNormalArguments == false)
            {
                for (int i = 0; i < argsResults.Count; i++)
                {
                    var pi = _function.ParametersInfo.GetParameterInfo(i);
                    if (EnumUtil.HasFlag(pi, FunctionParameterInformation.AdjustParameterAddress) && argsResults[i].Address != null)
                    {
                        _function.GetNewParameterAddress(argsResults, i, ref _dependencyAddresses);
                    }
                }
                return _dependencyAddresses != null;
            }
            return false;
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

                var compiler = Context.FunctionCompilerFactory.Create(_function, Context);
                var result = compiler.Compile(_args ?? Enumerable.Empty<CompileResult>(), Context);
                
                if (_negate != 0)
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
            while (i < _arguments.Count && _function.ParametersInfo.GetParameterInfo(i) != type) i++;
            if(i < _arguments.Count ) return _arguments[i];
            return -1;
        }
        internal override Expression CloneWithOffset(int row, int col)
        {
            if(_function==null || _function.ParametersInfo.HasNormalArguments)
            {
                return this;
            }
            return new FunctionExpression(Context) 
            {   
                _arguments = _arguments, 
                _function = _function, 
                _startPos = _startPos, 
                _endPos = _endPos
            };
        }

        internal string GetExpressionKey(RpnFormula f)
        {
            var key = new StringBuilder();
            for (int i=_startPos;i<_endPos;i++)
            {
                if (f._expressions.TryGetValue(i, out Expression e ))
                {
                    if(e.ExpressionType==ExpressionType.Function)
                    {
                        var fe = (FunctionExpression)e;
                        if (fe._function != null && fe._function.IsVolatile) return null;
                        key.Append(f._tokens[i].Value);
                    }
                    else
                    {
                        if(e.ExpressionType == ExpressionType.CellAddress)
                        {
                            var fa = e.GetAddress();
                            var adr = ExcelCellBase.GetAddress(fa.FromRow, fa.FromCol, fa.ToRow, fa.ToCol);
                            key.Append(adr);
                        }
                        else if(e.ExpressionType == ExpressionType.NameValue)
                        {
                            var ne = (NamedValueExpression)e;
                            if(ne._name!=null && ne.IsRelative)
                            {
                                key.Append($"{f._tokens[i].Value},{f._ws?.IndexInList},{f._row},{f._column}");
                            }
                            else
                            {
                                key.Append(f._tokens[i].Value);
                            }
                        }
                        else
                        {
                            var fa = e.GetAddress();
                            if(fa==null)
                            {
                                key.Append(f._tokens[i].Value);
                            }
                            else
                            {
                                var adr = fa.Address;
                                key.Append(adr);
                            }
                        }
                    }
                }
                else
                {
                    key.Append(f._tokens[i].Value);
                }
            }
            return key.ToString();
        }

        internal bool NeedsCheckAddressAdjustment()
        {
            if(_function.ParametersInfo.HasNormalArguments==false)
            {
                
            }
            return false;
        }

        private ExpressionStatus _status= ExpressionStatus.NoSet;
        internal override ExpressionStatus Status
        {
            get
            {
                if(_status == ExpressionStatus.NoSet)
                {
                    _status = ExpressionStatus.CanCompile;
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
