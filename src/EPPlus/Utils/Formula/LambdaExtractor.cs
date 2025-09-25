using OfficeOpenXml.FormulaParsing.FormulaExpressions;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Utils.Formula
{
    internal static class LambdaExtractor
    {
        public static ExtractedLambdaTokens GetLambdaTokens(string formula)
        {
            
            var tokens = SourceCodeTokenizer.Default.Tokenize(formula);
            var tokenIx = 0;
            var variableTokens = new List<Token>();
            var lambdaTokens = new List<Token>();
            bool collect = false;
            while(tokenIx < tokens.Count - 1)
            {
                var t = tokens[tokenIx];
                if(collect)
                {
                    lambdaTokens.Add(t);
                }
                else if(t.TokenType == TokenType.CommaLambda)
                {
                    collect = true;
                    tokenIx++;
                    continue;
                }
                else if(t.TokenType == TokenType.ParameterVariableDeclaration)
                {
                    variableTokens.Add(t);
                }
                tokenIx++;
            }
            var lambdaFormula = new StringBuilder();
            foreach (var tkn in lambdaTokens)
            {
                lambdaFormula.Append(tkn);
            }
            var tokensRpn = FormulaExecutor.CreateRPNTokens(lambdaTokens);
            var result = new ExtractedLambdaTokens();
            result.LambdaTokens = tokensRpn;
            result.VariableTokens = variableTokens;
            result.LambdaFormula = lambdaFormula.ToString();
            return result;
        }
    }
}
