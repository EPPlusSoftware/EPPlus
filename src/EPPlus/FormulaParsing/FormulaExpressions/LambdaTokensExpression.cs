/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  06/03/2025         EPPlus Software AB       Initial release EPPlus 8
 *************************************************************************************************/
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml.FormulaParsing.FormulaExpressions.VariableStorage;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;

namespace OfficeOpenXml.FormulaParsing.FormulaExpressions
{
    internal class LambdaTokensExpression : Expression
    {
        internal LambdaTokensExpression(ParsingContext ctx, VariableStorageScope scope) : base(ctx)
        {
            _scope = scope;
        }

        private List<Token> _tokens;
        private VariableStorageScope _scope;

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
            var result = new LambdaTokensResult(_tokens, _scope);
            return new CompileResult(result, DataType.LambdaTokens);
        }

        public override Expression Negate()
        {
            return this;
        }
    }
}
