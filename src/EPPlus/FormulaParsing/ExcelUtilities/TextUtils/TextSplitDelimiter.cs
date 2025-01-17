/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  12/13/2024         EPPlus Software AB           EPPlus 7.5
 *************************************************************************************************/
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.ExcelUtilities.TextUtils
{
    internal class TextSplitDelimiter
    {
        public TextSplitDelimiter(string delimiter, bool ignoreCase)
        {
            if(!string.IsNullOrEmpty(delimiter))
            {
                delimiter = delimiter.Replace("\"\"", "\"");
            }
            _delimiter = delimiter.ToCharArray();
            _ignoreCase = ignoreCase;
        }

        private readonly char[] _delimiter;
        private readonly bool _ignoreCase;
        private int _testIx = 0;

        public int DelimiterLength
        {
            get
            {
                return _delimiter?.Length ?? 0;
            }
        }
        public bool Test(char c)
        {
            if (_testIx >= DelimiterLength)
            {
                _testIx = 0;
                return true;
            }
            var dChar = _delimiter[_testIx];
            var isMatch = dChar == c || (_ignoreCase && char.ToUpper(dChar) == char.ToUpper(c));
            if(isMatch)
            {
                _testIx++;
            }
            if(_testIx == _delimiter.Length && _testIx > 0 && isMatch)
            {
                _testIx = 0;
                return true;
            }
            else if(!isMatch)
            {
                _testIx = 0;
            }
            return false;
        }
        

    }
}
