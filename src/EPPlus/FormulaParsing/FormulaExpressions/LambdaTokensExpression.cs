/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  08/20/2025         EPPlus Software AB       Initial release EPPlus 8.2
 *************************************************************************************************/
using OfficeOpenXml.FormulaParsing.FormulaExpressions.VariableStorage;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
using System.Collections.Generic;
using System.Diagnostics;

namespace OfficeOpenXml.FormulaParsing.FormulaExpressions
{
    /// <summary>
    /// An expression that holds the tokens of the Lambda expression, i.e. the tokens that 
    /// comes after the CommaLambda token. The <see cref="CompileResult"/> returned from the 
    /// <see cref="LambdaTokensExpression.Compile"/> function of this expression will have 
    /// an instance <see cref="LambdaTokensResult"/> as value with <see cref="DataType.LambdaTokens"/>.
    /// </summary>
    [DebuggerDisplay("LambdaTokensExpression - Count: {Tokens?.Count}")]
    internal class LambdaTokensExpression : Expression
    {
        internal LambdaTokensExpression(ParsingContext ctx, int variableScopeId) : base(ctx)
        {
            _scope = ctx.VariableStorage.GetById(variableScopeId);
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
