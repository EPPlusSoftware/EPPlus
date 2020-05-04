using System;
using System.Collections.Generic;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Engineering
{
    public static class BinaryHelper
    {
        public static bool TryParseBinaryToDecimal(string number, int raiseToBase, out int result)
        {
            var chars = number.ToCharArray();
            var isNegative = chars[0] == '1';
            var negativeUsed = false;
            result = 0;
            for (var x = 1; x < 10; x++)
            {
                var c = chars[x];
                var current = 0;
                if (c != '0' && c != '1') return false;
                if (x == 9)
                {
                    current = c == '1' ? 1 : 0;
                    if (isNegative && !negativeUsed) current *= -1;
                }
                else if (c == '1')
                {
                    current = (int)System.Math.Pow(raiseToBase, 9 - x);
                    if (isNegative && !negativeUsed) current *= -1;
                    negativeUsed = true;
                }
                result += current;
            }
            return true;
        }

        public static string EnsureLength(string input, int length, string padWith = "")
        {
            if (input == null) input = string.Empty;
            if(input.Length < length && !string.IsNullOrEmpty(padWith))
            {
                while(input.Length < length)
                {
                    input = padWith + input;
                }
            }
            else if(input.Length > length)
            {
                input = input.Substring(input.Length - length);
            }
            return input;
        }
    }
}
