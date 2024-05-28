using OfficeOpenXml.FormulaParsing.Excel.Functions.MathFunctions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.FormulaExpressions
{
    internal class LambdaCalculationExpression : Expression
    {
        public LambdaCalculationExpression(CompileResult cr, ParsingContext context) : base(context)
        {
            _compileResult = cr;
        }

        private readonly CompileResult _compileResult;

        internal override ExpressionType ExpressionType => ExpressionType.LambdaCalculation;

        internal override ExpressionStatus Status
        {
            get;
            set;
        } = ExpressionStatus.CanCompile;

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
    }
}
