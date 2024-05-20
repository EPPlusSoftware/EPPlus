/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  05/14/2024         EPPlus Software AB       Initial release EPPlus 7.3
 *************************************************************************************************/
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using OfficeOpenXml.FormulaParsing.FormulaExpressions;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;

namespace OfficeOpenXml.FormulaParsing.ExcelUtilities
{
    internal static class VariableParameterHelper
    {
#if (!NET35)
        [MethodImpl(MethodImplOptions.AggressiveInlining)]
#endif
        internal static bool IsVariableParameterFunction(string funcName)
        {
            if(string.IsNullOrEmpty(funcName)) return false;
            switch(funcName.ToLower())
            {
                case "_xlfn.let":
                case "let":
                    return true;
                default:
                    return false;
            }

        }
        private static bool IsVariableArg(string funcName, int argIndex, int argCount)
        {
            if(string.IsNullOrEmpty(funcName)) return false;
            funcName = funcName.Replace("_xlfn.", string.Empty);
            switch(funcName.ToLower())
            {
                case "let":
                    return argIndex % 2 == 0 && argIndex < argCount - 1;
                default:
                    return false;
            }
            
        }
        internal static void ProcessVariableArguments(IList<Token> tokens, int startIndex, string funcName)
        {
            var commaIndexes = new List<int>();
            var openParenthesis = 0;
            // 1. loop through the tokens and collect indexes of the commas until the end of the function args.
            for (var cIx = startIndex + 1; !(tokens[cIx].TokenType == TokenType.ClosingParenthesis && openParenthesis == 1); cIx++)
            {
                var t = tokens[cIx];
                if(t.TokenType == TokenType.OpeningParenthesis)
                {
                    openParenthesis++;
                }
                else if(t.TokenType == TokenType.ClosingParenthesis)
                {
                    openParenthesis--;
                }
                if (tokens[cIx].TokenType == TokenType.Comma && openParenthesis == 1)
                {
                    commaIndexes.Add(cIx);
                }
            }
            var variableNames = new Dictionary<string, int>();
            // 2. Process the arguments and look for variables excluding the last arg
            //    which is not declarations of variables but the calculation.
            for (var argIndex = 0; argIndex < commaIndexes.Count; argIndex++)
            {
                if (IsVariableArg(funcName, argIndex, commaIndexes.Count))
                {
                    var ix = commaIndexes[argIndex];
                    var variableName = ProcessVariableToken(tokens[ix - 1].Value);
                    if (!variableNames.ContainsKey(variableName))
                    {
                        variableNames.Add(variableName, argIndex);
                    }
                    tokens[ix - 1] = new Token(variableName, TokenType.ParameterVariable);
                }
                else
                {
                    var ix = commaIndexes[argIndex - 1] + 1;
                    while (tokens[ix].TokenType != TokenType.Comma || tokens[ix].TokenType == TokenType.Function)
                    {
                        var variableName = ProcessVariableToken(tokens[ix].Value);
                        if (variableNames.ContainsKey(variableName))
                        {
                            // if the declaration of the variable is using the variable a #NAME error should be returned.
                            if (variableNames[variableName] == argIndex - 1)
                            {
                                tokens[ix] = new Token(ExcelErrorValue.Create(eErrorType.Name).ToString(), TokenType.NameError);
                            }
                            else
                            {
                                tokens[ix] = new Token(variableName, TokenType.ParameterVariable);
                            }
                        }
                        ix++;
                    }
                }
            }
            // 3. Process variable names in the last argument (the calculation)
            openParenthesis = 1;
            for (var lastArgIx = commaIndexes.Last(); tokens[lastArgIx].TokenType != TokenType.ClosingParenthesis && openParenthesis == 1; lastArgIx++)
            {
                var candidate = tokens[lastArgIx];
                if (candidate.TokenType == TokenType.OpeningParenthesis)
                {
                    openParenthesis++;
                }
                else if (candidate.TokenType == TokenType.ClosingParenthesis)
                {
                    openParenthesis--;
                }
                // unresolved variable tokens will be interpreted as names
                if (candidate.TokenType == TokenType.NameValue)
                {
                    var candidateVariableName = ProcessVariableToken(candidate.Value);
                    if (variableNames.ContainsKey(candidateVariableName))
                    {
                        tokens[lastArgIx] = new Token(candidateVariableName, TokenType.ParameterVariable);
                    }
                }
            }
        }

        private static string ProcessVariableToken(string variableToken)
        {
            if (!string.IsNullOrEmpty(variableToken) && !variableToken.Trim().ToLower().Contains("_xlpm."))
            {
                return $"_xlpm.{variableToken.Trim()}";
            }
            return variableToken;
        }
    }
}
