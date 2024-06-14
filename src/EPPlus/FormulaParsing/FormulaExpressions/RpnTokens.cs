/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  06/12/2024         EPPlus Software AB       Initial release EPPlus 7.3
 *************************************************************************************************/
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;

namespace OfficeOpenXml.FormulaParsing.FormulaExpressions
{
    internal class RpnTokens : IEnumerable<Token>
    {
        public List<Token> Tokens { get; set; }

        public Token this[int index]
        {
            get => Tokens[index];
            set => Tokens[index] = value;
        }

        public int Count => Tokens.Count;

        public IEnumerator<Token> GetEnumerator()
        {
            return Tokens.GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return Tokens.GetEnumerator();
        }

        internal RpnTokens Clone()
        {
            var tokens = new List<Token>();
            foreach (var token in Tokens)
            {
                tokens.Add(new Token(token.Value, token.TokenType));
            }
            return new RpnTokens
            {
                Tokens = tokens
            };
        }

        internal Dictionary<int, int> LambdaRefs { get; set; }

        internal bool HasLambdaRefs
        {
            get
            {
                if (LambdaRefs == null) return false;
                return LambdaRefs.Count > 0;
            }
        }
    }
}
