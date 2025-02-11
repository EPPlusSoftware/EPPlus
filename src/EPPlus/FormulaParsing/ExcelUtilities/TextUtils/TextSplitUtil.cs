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
    internal class TextSplitUtil
    {

        public TextSplitUtil(string text)
        {
            _text = text;
            if (string.IsNullOrEmpty(text))
            {
                _textChars = [];
            }
            _textChars = text.ToCharArray();
        }

        private readonly string _text;
        private readonly char[] _textChars;

        private static bool EqualChars(char a, char b, bool ignoreCase)
        {
            if (!ignoreCase) return a == b;
            return char.ToUpper(a) == char.ToUpper(b);
        }

        public DelimiterInfo[] GetDelimiterPostions(string delimiter, bool ignoreCase = false, bool matchEnd = false)
        {
            return GetDelimiterPostions(new List<string> { delimiter }, ignoreCase);
        }

        public DelimiterInfo[] GetDelimiterPostions(List<string> delimiterInput, bool ignoreCase = false)
        {
            if (delimiterInput.Count == 0) return new DelimiterInfo[0];
            var delimiters = new List<TextSplitDelimiter>();
            foreach(var delStr in delimiterInput)
            {
                delimiters.Add(new TextSplitDelimiter(delStr, ignoreCase));
            }
            var delPostitions = new List<DelimiterInfo>();
            for (var cIx = 0; cIx < _textChars.Length; cIx++)
            {
                var c = _textChars[cIx];
                foreach(var delimiter in delimiters)
                {
                    if(delimiter.Test(c))
                    {
                        delPostitions.Add(new DelimiterInfo(delimiter.DelimiterLength, cIx - delimiter.DelimiterLength + 1));
                    }
                }
            }
            return delPostitions.ToArray();
        }



        private string GetText(bool getBefore, List<string> delimiters, int instanceNumber, bool ignoreCase, bool matchEnd, out bool isOutOfRange, out int? matchIndex)
        {
            isOutOfRange = false;
            var delimiterOccurances = GetDelimiterPostions(delimiters, ignoreCase);
            var inst = instanceNumber < 0 ? instanceNumber * -1 : instanceNumber;
            if (inst - 1 == delimiterOccurances.Length && matchEnd)
            {
                matchIndex = _text.Length;
                return _text;
            }
            if (delimiterOccurances.Length > 0 && instanceNumber < 0)
            {
                instanceNumber *= -1;
                Array.Reverse(delimiterOccurances);
            }
            var selectedPosIx = instanceNumber - 1;
            if (selectedPosIx < 0 || selectedPosIx >= delimiterOccurances.Length)
            {
                isOutOfRange = true;
                matchIndex = null;
                return string.Empty;
            }
            var del = delimiterOccurances[selectedPosIx];
            if (del.Position >= _text.Length)
            {
                del.Position = _text.Length;
            }
            matchIndex = del.Position;
            return getBefore ? _text.Substring(0, del.Position) : _text.Substring(del.Position + del.Length);
        }

        public string GetTextBefore(string delimiter, int instanceNumber, bool ignoreCase, bool matchEnd, out bool isOutOfRange, out int? matchIndex)
        {
            return GetTextBefore(new List<string> { delimiter }, instanceNumber, ignoreCase, matchEnd, out isOutOfRange, out matchIndex);
        }

        public string GetTextBefore(List<string> delimiters, int instanceNumber, bool ignoreCase, bool matchEnd, out bool isOutOfRange, out int? matchIndex)
        {
            return GetText(true, delimiters, instanceNumber, ignoreCase, matchEnd, out isOutOfRange, out matchIndex);
        }

        public string GetTextAfter(string delimiter, int instanceNumber, bool ignoreCase, bool matchEnd, out bool isOutOfRange, out int? matchIndex)
        {
            return GetTextAfter(new List<string> { delimiter }, instanceNumber, ignoreCase, matchEnd, out isOutOfRange, out matchIndex);
        }

        public string GetTextAfter(List<string> delimiters, int instanceNumber, bool ignoreCase, bool matchEnd, out bool isOutOfRange, out int? matchIndex)
        {
            return GetText(false, delimiters, instanceNumber, ignoreCase, matchEnd, out isOutOfRange, out matchIndex);
        }
    }
}
