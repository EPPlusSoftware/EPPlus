using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;

namespace OfficeOpenXml.FormulaParsing.FormulaExpressions
{
    internal class LambdaCalculationExpression : Expression
    {
        internal LambdaCalculationExpression(List<Token> tokens, ParsingContext ctx) : base(ctx)
        {
            _tokens = tokens;
        }

        private readonly List<Token> _tokens;

        public List<Token> Tokens => _tokens;

        internal override ExpressionType ExpressionType => ExpressionType.LambdaCalculation;

        internal override ExpressionStatus Status
        {
            get;
            set;
        } = ExpressionStatus.CanCompile;

        public override CompileResult Compile()
        {
            return new CompileResult(_tokens, DataType.LambdaCalculation);
        }

        public override Expression Negate()
        {
            return this;
        }
    }
}
