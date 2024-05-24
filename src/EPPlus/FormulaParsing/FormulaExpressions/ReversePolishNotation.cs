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

namespace OfficeOpenXml.FormulaParsing.FormulaExpressions
{
    internal static class ReversePolishNotation
    {
        internal static List<Token> CreateRPNTokens(IList<Token> tokens)
        {
            var bracketCount = 0;
            var operators = OperatorsDict.Instance;
            Stack<Token> operatorStack = new Stack<Token>();
            var expressions = new List<Token>();
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
                                expressions.Add(o);
                                if (operatorStack.Count == 0) throw new InvalidOperationException("No closing parenthesis");
                                o = operatorStack.Pop();
                            }
                            if (operatorStack.Count > 0 && operatorStack.Peek().TokenType == TokenType.Function)
                            {
                                expressions.Add(operatorStack.Pop());
                            }
                        }
                        break;
                    case TokenType.Operator:
                    case TokenType.Negator:
                        if (token.TokenType == TokenType.Operator && i > 0 && i < tokens.Count - 2 && token.Value == ":" && tokens[i - 1].Value == "]" && tokens[i + 1].Value == "[")
                        {
                            expressions.Add(token);
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
                                expressions.Add(operatorStack.Pop());
                                if (operatorStack.Count == 0) break;
                                o2 = operatorStack.Peek();
                            }
                        }
                        operatorStack.Push(token);
                        break;

                    case TokenType.Function:
                        expressions.Add(new Token(token.Value, TokenType.StartFunctionArguments));
                        operatorStack.Push(token);
                        break;
                    case TokenType.Comma:
                    case TokenType.CommaLambda:
                        if (operatorStack.Count > 0 && bracketCount == 0) //If inside a table 
                        {
                            var op = operatorStack.Peek().TokenType;
                            while (op == TokenType.Operator || op == TokenType.Negator)
                            {
                                expressions.Add(operatorStack.Pop());
                                if (operatorStack.Count == 0) break;
                                op = operatorStack.Peek().TokenType;
                            }
                        }
                        expressions.Add(token);
                        break;
                    case TokenType.OpeningBracket:
                        bracketCount++;
                        expressions.Add(token);
                        break;
                    case TokenType.ClosingBracket:
                        bracketCount--;
                        expressions.Add(token);
                        break;
                    default:
                        expressions.Add(token);
                        break;
                }

            }

            while (operatorStack.Count > 0)
            {
                expressions.Add(operatorStack.Pop());
            }

            return expressions;
        }
    }
}
