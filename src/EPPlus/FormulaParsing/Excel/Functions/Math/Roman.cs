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
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using System;
using System.Collections.Generic;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Math
{
    internal class Roman : ExcelFunction
    {
        private readonly RomanNumber One = new RomanNumber(1, "I");
        private readonly RomanNumber Five = new RomanNumber(5, "V");
        private readonly RomanNumber Ten = new RomanNumber(10, "X");
        private readonly RomanNumber Fifty = new RomanNumber(50, "L");
        private readonly RomanNumber OneHundred = new RomanNumber(100, "C");
        private readonly RomanNumber FiveHundred = new RomanNumber(500, "D");
        private readonly RomanNumber Thousand = new RomanNumber(1000, "M");

        class RomanNumber
        {
            public RomanNumber(int number, string letter)
            {
                Number = number;
                Letter = letter;
            }
            public int Number { get; set; }

            public string Letter { get; set; }
        }

        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            ValidateArguments(arguments, 1);
            var number = ArgToInt(arguments, 0);
            if (number < 0 || number > 3999) return CreateResult(eErrorType.Value);
            var result = new StringBuilder();
            Apply(ref number, Thousand, result);
            Apply(ref number, 900, "CM", result);
            Apply(ref number, FiveHundred, OneHundred, result);
            Apply(ref number, 400, "CD", result);
            Apply(ref number, OneHundred, result);
            Apply(ref number, 90, "XC", result);
            Apply(ref number, Fifty, Ten, result);
            Apply(ref number, 40, "XL", result);
            Apply(ref number, Ten, result);
            Apply(ref number, 9, "IX", result);
            Apply(ref number, Five, One, result);
            Apply(ref number, 4, "IV", result);
            Apply(ref number, One, result);
            return CreateResult(result.ToString(), DataType.String);
        }

        private void Apply(ref int number, RomanNumber roman, StringBuilder result)
        {
            if (number >= roman.Number)
            {
                var limit = number / roman.Number;
                for (var x = 0; x < limit; x++)
                {
                    result.Append(roman.Letter);
                    number -= roman.Number;
                }
            }
        }

        private void Apply(ref int number, RomanNumber roman, RomanNumber lowerRoman, StringBuilder result)
        {
            if(number >= roman.Number)
            {
                result.Append(roman.Letter);
                number -= roman.Number;
                var limit = number / lowerRoman.Number;
                for(var x = 0; x < (number / lowerRoman.Number); x++)
                {
                    result.Append(lowerRoman.Letter);
                    number -= lowerRoman.Number;
                }
            }
        }

        private void Apply(ref int number, int limit, string letters, StringBuilder result)
        {
            if(number >= limit)
            {
                result.Append(letters);
                number -= limit;
            }
        }

    }
}
