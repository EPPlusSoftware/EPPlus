using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;

namespace OfficeOpenXml.FormulaParsing.FormulaExpressions
{
    internal class LambdaTokensExpression : Expression
    {
        internal LambdaTokensExpression(ParsingContext ctx) : base(ctx)
        {

        }

        private List<Token> _tokens;

        public List<Token> Tokens => _tokens;

        internal void AddLambdaToken(Token token)
        {
            _tokens ??= [];
            _tokens.Add(token);
        }

        internal override ExpressionType ExpressionType => ExpressionType.LambdaCalculation;

        internal override ExpressionStatus Status
        {
            get;
            set;
        } = ExpressionStatus.CanCompile;

        public override CompileResult Compile()
        {
            return new CompileResult(_tokens, DataType.LambdaTokens);
        }

        public override Expression Negate()
        {
            return this;
        }
    }
}
