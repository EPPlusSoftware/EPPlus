/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  06/12/2024         EPPlus Software AB       Initial release EPPlus 7.3
 *************************************************************************************************/
using OfficeOpenXml.FormulaParsing.Excel.Functions.Logical;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.FormulaExpressions
{
    internal class LambdaInvokeExpression : FunctionExpression
    {
        internal LambdaInvokeExpression(LambdaCalculationExpression exp, ParsingContext ctx, int pos) : base("LAMBDAINVOKE", ctx, pos)
        {
            _calculationExpression = exp;
            this._function = new Lambda();
        }

        private bool _negate = false;
        private readonly LambdaCalculationExpression _calculationExpression;
        private readonly List<CompileResult> _lambdaArguments = new List<CompileResult>();

        internal override ExpressionType ExpressionType => ExpressionType.LambdaInvoke;

        internal void AddArgument(CompileResult compileResult)
        {
            _lambdaArguments.Add(compileResult);
        }

        public override CompileResult Compile()
        {
            if(_calculationExpression == null)
            {
                return CompileResult.GetErrorResult(eErrorType.Value);
            }
            var cr = _calculationExpression.Compile();
            if(cr.DataType != DataType.LambdaCalculation || cr.Result == null)
            {
                return CompileResult.GetErrorResult(eErrorType.Value);
            }
            var calculator = cr.Result as LambdaCalculator;
            calculator.BeginCalculation();
            for (var i = 0; i < _lambdaArguments.Count; i++)
            {
                var arg = _lambdaArguments[i];
                calculator.SetVariableValue(i, arg.Result, arg.DataType);
            }
            return calculator.Execute(Context);
        }

        public override Expression Negate()
        {
            _negate = !_negate;
            return this;
        }

        private ExpressionStatus _status = ExpressionStatus.NoSet;
        internal override ExpressionStatus Status
        {
            get
            {
                if (_status == ExpressionStatus.NoSet)
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
