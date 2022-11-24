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
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
using System.Collections;
using System.Collections.Generic;

namespace OfficeOpenXml
{
    internal class TokenOffsetCollection : IEnumerable<Token>, IEnumerator<Token>
    {
        private string _currentWs;
        private int _rowOffset;
        private int _columnOffset;
        private List<Token> _tokens;
        private int _index=0;
        private Dictionary<int, ShiftableAddress> TokenAddresses { get; set; }
        public TokenOffsetCollection(string currentWs, List<Token> tokens)
        {
            _currentWs = currentWs;
            _tokens = tokens;
            TokenAddresses=new Dictionary<int, ShiftableAddress>();
            for (var x = 0; x < _tokens.Count; x++)
            {
                var t = tokens[x];
                if (t.TokenType == TokenType.ExcelAddress)
                {
                    TokenAddresses.Add(x,new ShiftableAddress(t.Value));
                }
            }
        }

        public void SetOffset(int rowOffset, int columnOffset)
        {
            _rowOffset = rowOffset;
            _columnOffset = columnOffset;
        }

        public Token Current
        {
            get
            {
                var t = _tokens[_index];
                if(t.TokenType==TokenType.ExcelAddress && (_rowOffset != 0 || _columnOffset != 0)) 
                {
                    var sa = TokenAddresses[_index];
                    var newAddress = sa.GetOffsetAddress(_rowOffset, _columnOffset);
                    return new Token(newAddress, TokenType.ExcelAddress);
                }
                else
                {
                    return t;
                }
            }
        }

        object IEnumerator.Current => throw new System.NotImplementedException();

        public void Dispose()
        {
            
        }

        public IEnumerator<Token> GetEnumerator()
        {
            return this;
        }

        public bool MoveNext()
        {
            if (_index == _tokens.Count - 1)
            {
                _index = -1;
                return false;
            }
            _index++;
            return true;
        }

        public void Reset()
        {
            _index = -1;
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return this;
        }
    }
}