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
    internal class VariableParameterHelper
    {
        public VariableParameterHelper(IList<Token> tokens, List<int> funcPositions)
        {
            _tokens = tokens;
            _funcPositions = funcPositions;
            _functions = new List<VariableFunction>();
            foreach(var pos in  funcPositions)
            {
                var f = new VariableFunction(_functions)
                {
                    Name = tokens[pos].Value,
                    Start = pos
                };
                _functions.Add(f);
            }
        }

        private readonly IList<Token> _tokens;
        private readonly List<int> _funcPositions;
        private readonly List<VariableFunction> _functions;

        private class VariableFunction
        {
            public VariableFunction(List<VariableFunction> functions)
            {
                _functions = functions;
            }

            private readonly List<VariableFunction> _functions;

            public string Name { get; set; }

            public int Start { get; set; }

            public int? End { get; set; }

            public Dictionary<string, int> Variables { get; set; } = new Dictionary<string, int>();

            public bool IsGlobalVariable(string name)
            {
                if (Variables.ContainsKey(name)) return true;
                var parentFunctions = _functions.Where(f => f.Start < Start && (f.End > Start || f.End == null));
                foreach(var parentFunc in parentFunctions)
                {
                    if(parentFunc.Variables.ContainsKey(name)) return true;
                }
                return false;
            }

            public bool IsLocalVariable(string name)
            {
                return Variables.ContainsKey(name);
            }
        }

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

        internal void Process()
        {
            foreach(var func in _functions)
            {
                ProcessVariableArguments(func);
            }
        }

        private void ProcessVariableArguments(VariableFunction func)
        {
            var commaIndexes = new List<int>();
            var openParenthesis = 0;
            // 1. loop through the tokens and collect indexes of the commas until the end of the function args.
            for (var cIx = func.Start + 1; !(_tokens[cIx].TokenType == TokenType.ClosingParenthesis && openParenthesis == 1); cIx++)
            {
                var t = _tokens[cIx];
                if(t.TokenType == TokenType.OpeningParenthesis)
                {
                    openParenthesis++;
                }
                else if(t.TokenType == TokenType.ClosingParenthesis)
                {
                    openParenthesis--;
                }
                if (_tokens[cIx].TokenType == TokenType.Comma && openParenthesis == 1)
                {
                    commaIndexes.Add(cIx);
                }
            }
            // 2. Process the arguments and look for variables excluding the last arg
            //    which is not declarations of variables but the calculation.
            var lastDeclaredVariable = string.Empty;
            for (var argIndex = 0; argIndex < commaIndexes.Count; argIndex++)
            {
                if (IsVariableArg(func.Name, argIndex, commaIndexes.Count))
                {
                    var ix = commaIndexes[argIndex];
                    var variableName = ProcessVariableToken(_tokens[ix - 1].Value);
                    lastDeclaredVariable = variableName;
                    if (!func.IsLocalVariable(variableName))
                    {
                        func.Variables.Add(variableName, argIndex);
                    }
                    _tokens[ix - 1] = new Token(variableName, TokenType.ParameterVariableDeclaration);
                }
                else
                {
                    var ix = commaIndexes[argIndex - 1] + 1;
                    while (_tokens[ix].TokenType != TokenType.Comma || _tokens[ix].TokenType == TokenType.Function)
                    {
                        // Variables will per default have token type NameValue. Don't process tokens that have other token types.
                        var t = _tokens[ix];
                        var variableName = ProcessVariableToken(_tokens[ix].Value);
                        if (func.IsGlobalVariable(variableName))
                        {
                            if (variableName == lastDeclaredVariable)
                            {
                                _tokens[ix] = new Token(variableName, TokenType.NameError);
                            }
                            else
                            {
                                _tokens[ix] = new Token(variableName, TokenType.ParameterVariable);
                            }
                        }


                        ix++;
                    }
                    lastDeclaredVariable = string.Empty;
                }
            }
            // 3. Process variable names in the last argument (the calculation)
            openParenthesis = 1;
            int lastArgIx;
            for (lastArgIx = commaIndexes.Last(); _tokens[lastArgIx].TokenType != TokenType.ClosingParenthesis && openParenthesis == 1; lastArgIx++)
            {
                var candidate = _tokens[lastArgIx];
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
                    if (func.IsGlobalVariable(candidateVariableName))
                    {
                        _tokens[lastArgIx] = new Token(candidateVariableName, TokenType.ParameterVariable);
                    }
                }
            }
            func.End = lastArgIx;
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
