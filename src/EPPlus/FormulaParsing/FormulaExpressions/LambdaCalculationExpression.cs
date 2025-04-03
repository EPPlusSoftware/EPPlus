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
using OfficeOpenXml.FormulaParsing.Excel.Functions.MathFunctions;
using OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.FormulaExpressions
{
    [DebuggerDisplay("LambdaCalculationExpression - _compileResult: {GetResultInfo()}")]
    internal class LambdaCalculationExpression : Expression
    {
        public LambdaCalculationExpression(CompileResult cr, ParsingContext context) : base(context)
        {
            _compileResult = cr;
            Id = Guid.NewGuid();
        }

        private readonly CompileResult _compileResult;

        internal string GetResultInfo()
        {
            if (_compileResult == null) return "<null>";
            if(_compileResult.Result is LambdaCalculator calc)
            {
                return calc.GetDebugInfo();
            }
            return _compileResult.ToString();   
        }

        internal override ExpressionType ExpressionType => ExpressionType.LambdaCalculation;

        internal override ExpressionStatus Status
        {
            get;
            set;
        } = ExpressionStatus.CanCompile;

        public Guid Id { get; private set; }

        public override CompileResult Compile()
        {
            if(_compileResult.DataType != DataType.LambdaCalculation)
            {
                return CompileResult.GetErrorResult(eErrorType.Value);
            }
            return _compileResult;
        }

        public override Expression Negate()
        {
            throw new NotImplementedException();
        }

        public void SetVariable(int index, object val, DataType dt)
        {
            if(_compileResult.Result is LambdaCalculator calculator)
            {
                calculator.SetVariableValue(index, val, dt, Context);
            }
        }

        public int GetNumberOfVariables()
        {
            if(_compileResult.Result is LambdaCalculator lce)
            {
                return lce.NumberOfVariables;
            }
            return 0;
        }

        public void BeginCalculation()
        {
            if (_compileResult.Result is LambdaCalculator calculator)
            {
                calculator.BeginCalculation();
            }
        }
    }
}
