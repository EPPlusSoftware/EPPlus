/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  25/07/2023         EPPlus Software AB           EPPlus v7
 *************************************************************************************************/

using OfficeOpenXml.FormulaParsing.Excel.Functions.Metadata;
using OfficeOpenXml.FormulaParsing.FormulaExpressions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.MathFunctions
{

    [FunctionMetadata(
    Category = ExcelFunctionCategory.MathAndTrig,
    EPPlusVersion = "7.0",
    Description = "Returns an integer depicting its Roman equivalence.",
    SupportsArrays = true)]
    internal class Arabic : ExcelFunction
    {
        public static Dictionary<string, int> RomanToArabic = new Dictionary<string, int>()
        {
            {"I", 1 },
            {"V", 5 },
            {"X", 10 },
            {"L", 50 },
            {"C", 100 },
            {"D", 500 },
            {"M", 1000 }
        };

        public override int ArgumentMinLength => 1;

        public override ExcelFunctionArrayBehaviour ArrayBehaviour => ExcelFunctionArrayBehaviour.FirstArgCouldBeARange;

        public override CompileResult Execute(IList<FunctionArgument> arguments, ParsingContext context)
        {
            var romanString = ArgToString(arguments, 0);
            romanString = romanString.ToUpper().Trim();
            bool negativeRoman = false;

            foreach (char romanChar in romanString)
            {
                if (romanChar == '-') negativeRoman = true;
                if (!(RomanToArabic.ContainsKey(romanChar.ToString())) && romanChar != '-')
                {
                    return CreateResult(eErrorType.Value);
                }
            }
            if (negativeRoman) romanString = romanString.Remove(0, 1);
            var romanLength = romanString.Length;

            if (romanLength > 255) return CreateResult(eErrorType.Value);
            if (romanLength == 0) return CreateResult(0d, DataType.Decimal);

            var arabicResult = 0d;
            var currentTopValue = 0d;

            for (int romanIndex = romanLength - 1; romanIndex >= 0d; romanIndex--)
            {
                if (ArabicVal(romanIndex, romanString) >= currentTopValue)
                {
                    currentTopValue = ArabicVal(romanIndex, romanString);
                    arabicResult += currentTopValue;
                }
                else
                {
                    arabicResult -= ArabicVal(romanIndex, romanString);
                }
            }
            if (negativeRoman) arabicResult = -arabicResult;
            return CreateResult(arabicResult, DataType.Decimal);
        }

        public static double ArabicVal(int index, string txt)
        {
            //Returns arabic value from the dictionary
            return RomanToArabic[txt.Substring(index, 1)];
        }
    }
}
