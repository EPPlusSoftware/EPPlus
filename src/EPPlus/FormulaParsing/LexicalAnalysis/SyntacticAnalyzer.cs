/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  01/27/2020         EPPlus Software AB       Initial release EPPlus 5
 *************************************************************************************************/
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml.FormulaParsing.Exceptions;

namespace OfficeOpenXml.FormulaParsing.LexicalAnalysis
{
    /// <summary>
    /// Validates the state of parsed tokens and throws exceptions if they are not valid according to
    /// the following:
    /// - All opened and closed parenthesis must match
    /// - String must be closed
    /// - There must be no unrecognized tokens
    /// </summary>
    public class SyntacticAnalyzer : ISyntacticAnalyzer
    {
        private class AnalyzingContext
        {
            public int NumberOfOpenedParentheses { get; set; }
            public int NumberOfClosedParentheses { get; set; }
            public int OpenedStrings { get; set; }
            public int ClosedStrings { get; set; }
            public bool IsInString { get; set; }
        }

        /// <summary>
        /// Analyzes the parsed tokens.
        /// </summary>
        /// <param name="tokens"></param>
        public void Analyze(IEnumerable<Token> tokens)
        {
            var context = new AnalyzingContext();
            foreach (var token in tokens)
            {
                if (token.TokenTypeIsSet(TokenType.Unrecognized))
                {
                    throw new UnrecognizedTokenException(token);
                }
                EnsureParenthesesAreWellFormed(token, context);
                EnsureStringsAreWellFormed(token, context);
            }
            Validate(context);
        }

        private static void Validate(AnalyzingContext context)
        {
            if (context.NumberOfOpenedParentheses != context.NumberOfClosedParentheses)
            {
                throw new FormatException("Number of opened and closed parentheses does not match");
            }
            if (context.OpenedStrings != context.ClosedStrings)
            {
                throw new FormatException("Unterminated string");
            }
        }

        private void EnsureParenthesesAreWellFormed(Token token, AnalyzingContext context)
        {
            if (token.TokenTypeIsSet(TokenType.OpeningParenthesis))
            {
                context.NumberOfOpenedParentheses++;
            }
            else if (token.TokenTypeIsSet(TokenType.ClosingParenthesis))
            {
                context.NumberOfClosedParentheses++;
            }
        }

        private void EnsureStringsAreWellFormed(Token token, AnalyzingContext context)
        {
            if (!context.IsInString && token.TokenTypeIsSet(TokenType.String))
            {
                context.IsInString = true;
                context.OpenedStrings++;
            }
            else if (context.IsInString && token.TokenTypeIsSet(TokenType.String))
            {
                context.IsInString = false;
                context.ClosedStrings++;
            }
        }
    }
}
