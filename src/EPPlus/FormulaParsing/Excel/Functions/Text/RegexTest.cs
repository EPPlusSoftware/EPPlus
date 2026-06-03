using OfficeOpenXml.FormulaParsing.FormulaExpressions;
using OfficeOpenXml.FormulaParsing.Ranges;
using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Text
{
    internal class RegexTest : RegexFunctionBase
    {
        public override int ArgumentMinLength => 2;
        public override string NamespacePrefix => "_xlfn.";
        
        public override CompileResult Execute(IList<FunctionArgument> arguments, ParsingContext context)
        {
            bool textIsRange = arguments[0].IsExcelRange;
            bool patternIsRange = arguments[1].IsExcelRange;
            int caseSensitivity = arguments.Count > 2 ? ArgToInt(arguments, 2, 0) : 0;

            if (!textIsRange && !patternIsRange)
            {
                var text = arguments[0].Value?.ToString();
                var pattern = arguments[1].Value?.ToString();

                if (text == null || pattern == null)
                    return CreateResult(ExcelErrorValue.Create(eErrorType.NA), DataType.ExcelError);
                if (caseSensitivity > 1 || caseSensitivity < 0)
                    return CreateResult(ExcelErrorValue.Create(eErrorType.Value), DataType.ExcelError);

                return CreateResult(GetRegexTest(text, pattern, caseSensitivity), DataType.Boolean);
            }

            var texts = textIsRange ? arguments[0].ValueAsRangeInfo : null;
            var patterns = patternIsRange ? arguments[1].ValueAsRangeInfo : null;

            int textRows = texts != null ? texts.Size.NumberOfRows : 1;
            int textCols = texts != null ? texts.Size.NumberOfCols : 1;
            int patternRows = patterns != null ? patterns.Size.NumberOfRows : 1;
            int patternCols = patterns != null ? patterns.Size.NumberOfCols : 1;

            var nRows = ExpandedSize(textRows, patternRows);
            var nCols = ExpandedSize(textCols, patternCols);
                        
            var result = new InMemoryRange(nRows, nCols);

            for (int row = 0; row < nRows; row++)
            {
                for (int col = 0; col < nCols; col++)
                {
                    var textValue = GetValue(texts, arguments[0], textRows, textCols, row, col);
                    var patternValue = GetValue(patterns, arguments[1], patternRows, patternCols, row, col);

                    if (textValue == null || patternValue == null)
                        result.SetValue(row, col, ExcelErrorValue.Create(eErrorType.NA));
                    else if(caseSensitivity > 1 || caseSensitivity < 0)
                    {
                        result.SetValue(row, col, ExcelErrorValue.Create(eErrorType.Value));
                    }
                    else
                        result.SetValue(row, col, GetRegexTest(textValue, patternValue, caseSensitivity));
                }
            }

            return CreateDynamicArrayResult(result, DataType.ExcelRange);
        }

        private static bool GetRegexTest(string text, string pattern, int caseSensitive)
            => Regex.IsMatch(text, pattern, (RegexOptions)caseSensitive);
    }
}