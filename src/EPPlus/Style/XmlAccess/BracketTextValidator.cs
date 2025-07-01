/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  06/27/2025         EPPlus Software AB       Improved format handling
 *************************************************************************************************/
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Style.XmlAccess
{
    internal static class BracketTextValidator
    {
        private static HashSet<string> _colors = new HashSet<string>(StringComparer.InvariantCultureIgnoreCase) { "Black", "Green", "White", "Blue", "Magenta", "Yellow", "Cyan", "Red" };

        private static HashSet<char> _operators = new HashSet<char> { '>', '<', '=' };

        public static bool IsValid(string text)
        {
            if(string.IsNullOrEmpty(text)) return false;
            if (_colors.Contains(text)) return true;
            if (text.StartsWith("$")) return true;
            if(IsCondition(text)) return true;
            // indexed colors
            if (text.ToLowerInvariant().StartsWith("color") && int.TryParse(text.Substring(5), out int n) && n >= 1 && n <= 56) return true;
            return false;
        }

        public static bool IsCondition(string text)
        {

            if (string.IsNullOrEmpty(text)) return false;

            text = text.Replace(" ", "");

            int i = 0;
            if (text.StartsWith(">=")) i = 2;
            else if (text.StartsWith("<=")) i = 2;
            else if (text.StartsWith(">") || text.StartsWith("<") || text.StartsWith("=")) i = 1;
            else return false;

            if (i >= text.Length) return false;

            bool isValidNumber = false;
            bool hasDecimal = false;

            if (text[i] == '-') i++; // tillåt negativt tal

            for (; i < text.Length; i++)
            {
                char c = text[i];
                if (char.IsDigit(c))
                {
                    isValidNumber = true;
                }
                else if (c == '.' && !hasDecimal)
                {
                    hasDecimal = true;
                }
                else
                {
                    return false;
                }
            }

            return isValidNumber;

        }
    }
}
