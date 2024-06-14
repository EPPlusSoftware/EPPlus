/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  01/27/2024         EPPlus Software AB       Initial release EPPlus 7
 *************************************************************************************************/
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
using OfficeOpenXml.FormulaParsing.Excel.Operators;
using OfficeOpenXml.FormulaParsing.Utilities;

namespace OfficeOpenXml.FormulaParsing.FormulaExpressions
{
    internal static class ReversePolishNotation
    {
        internal static RpnTokens CreateRPNTokens(IList<Token> tokens)
        {
            var bracketCount = 0;
            var operators = OperatorsDict.Instance;
            Stack<Token> operatorStack = new Stack<Token>();
            Stack<int> lambdas = new Stack<int>();
            var rpnTokens = new List<Token>();
            var hasLambda = false;
            for (int i = 0; i < tokens.Count; i++)
            {
                Token token = tokens[i];
                switch (token.TokenType)
                {
                    case TokenType.OpeningParenthesis:
                        operatorStack.Push(token);
                        break;
                    case TokenType.ClosingParenthesis:
                        if (operatorStack.Count > 0)
                        {
                            var o = operatorStack.Pop();
                            while (o.TokenType != TokenType.OpeningParenthesis)
                            {
                                rpnTokens.Add(o);
                                if (operatorStack.Count == 0) throw new InvalidOperationException("No closing parenthesis");
                                o = operatorStack.Pop();
                            }
                            if (operatorStack.Count > 0 && operatorStack.Peek().TokenType == TokenType.Function)
                            {
                                rpnTokens.Add(operatorStack.Pop());
                            }
                        }
                        break;
                    case TokenType.Operator:
                    case TokenType.Negator:
                        if (token.TokenType == TokenType.Operator && i > 0 && i < tokens.Count - 2 && token.Value == ":" && tokens[i - 1].Value == "]" && tokens[i + 1].Value == "[")
                        {
                            rpnTokens.Add(token);
                            break;
                        }
                        if (operatorStack.Count > 0)

                        {
                            var o2 = operatorStack.Peek();
                            while ((o2.TokenType == TokenType.Operator && token.TokenType != TokenType.Negator &&
                                operators[o2.Value].Precedence <= operators[token.Value].Precedence)
                                ||
                                (o2.TokenType == TokenType.Negator &&
                                token.TokenType != TokenType.Negator &&
                                operators[token.Value].Precedence > Operator.PrecedenceColon))
                            {
                                rpnTokens.Add(operatorStack.Pop());
                                if (operatorStack.Count == 0) break;
                                o2 = operatorStack.Peek();
                            }
                        }
                        operatorStack.Push(token);
                        break;

                    case TokenType.Function:
                        rpnTokens.Add(new Token(token.Value, TokenType.StartFunctionArguments));
                        operatorStack.Push(token);
                        break;
                    case TokenType.Comma:
                    case TokenType.CommaLambda:
                        if (operatorStack.Count > 0 && bracketCount == 0) //If inside a table 
                        {
                            var op = operatorStack.Peek().TokenType;
                            while (op == TokenType.Operator || op == TokenType.Negator)
                            {
                                rpnTokens.Add(operatorStack.Pop());
                                if (operatorStack.Count == 0) break;
                                op = operatorStack.Peek().TokenType;
                            }
                        }
                        if(token.TokenType == TokenType.CommaLambda)
                        {
                            hasLambda = true;
                        }
                        rpnTokens.Add(token);
                        break;
                    case TokenType.OpeningBracket:
                        bracketCount++;
                        rpnTokens.Add(token);
                        break;
                    case TokenType.ClosingBracket:
                        bracketCount--;
                        rpnTokens.Add(token);
                        break;
                    default:
                        rpnTokens.Add(token);
                        break;
                }

            }

            while (operatorStack.Count > 0)
            {
                rpnTokens.Add(operatorStack.Pop());
            }
            var result = new RpnTokens
            {
                Tokens = rpnTokens
            };
            if (hasLambda)
            {
                ProcessLambda(result);
            }
            return result;
        }

        private static void ProcessLambda(RpnTokens rpnTokens)
        {
            var lambdaRefs = new Dictionary<int, int>();
            Stack<int> lStack = new Stack<int>();
            for(var i = 0; i < rpnTokens.Count; i++)
            {
                var token = rpnTokens[i];
                if(token.IsLambdaFunction())
                {
                    if(token.TokenType == TokenType.StartFunctionArguments)
                    {
                        lStack.Push(i);
                    }
                    else
                    {
                        lambdaRefs[lStack.Pop()] = i + 1;
                    }
                }
            }
            rpnTokens.LambdaRefs = lambdaRefs;

            // TODO: create a new list and loop through the existing tokens
            // move tokens from the Lambda invoke parenthesis to after the corresponding
            // variable and use the new Assign operator... /MA
            var TokenList = new List<Token>();
        }

        private static void MoveToken(ref List<Token> tokens, int fromPos, int toPos) 
        {
            var token = tokens[fromPos];
            tokens.Insert(toPos, token);
            tokens.RemoveAt(fromPos);

        }

        private static void InsertAssignOperator(ref List<Token> tokens, int pos)
        {
            tokens.Insert(pos, new Token(Operator.AssignIndicator, TokenType.Operator));
        }
    }
}
