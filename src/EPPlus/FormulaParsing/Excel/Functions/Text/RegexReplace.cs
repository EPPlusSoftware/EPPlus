using OfficeOpenXml.FormulaParsing.FormulaExpressions;
using OfficeOpenXml.FormulaParsing.Ranges;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Text.RegularExpressions;


namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Text
{
    internal class RegexReplace : RegexFunctionBase
    {
        public override int ArgumentMinLength => 3;

        public override string NamespacePrefix => "_xlfn.";

        public override CompileResult Execute(IList<FunctionArgument> arguments, ParsingContext context)
        {
            bool textIsRange = arguments[0].IsExcelRange;
            bool patternIsRange = arguments[1].IsExcelRange;
            bool replacementIsRange = arguments[2].IsExcelRange;

            int occurnance = arguments.Count > 3 ? ArgToInt(arguments, 3, 0) : 0;
            int caseSensitive = arguments.Count > 4 ? ArgToInt(arguments, 4, 0) : 0;

            if (!textIsRange && !patternIsRange && !replacementIsRange)
            {
                var text = arguments[0].Value?.ToString() ?? string.Empty;
                var pattern = arguments[1].Value?.ToString() ?? string.Empty;
                var replacement = arguments[2].Value?.ToString() ?? string.Empty;

                if (caseSensitive > 1 || caseSensitive < 0 || (text != null && pattern == string.Empty))
                    return CreateResult(ExcelErrorValue.Create(eErrorType.Value), DataType.ExcelError);
                var res = GetRegexReplaced(text, pattern, replacement, occurnance, caseSensitive);
                if (res == null)
                    return CreateResult(ExcelErrorValue.Create(eErrorType.Value), DataType.ExcelError);
                return CreateResult(res, DataType.String);
            }

            var texts = textIsRange ? arguments[0].ValueAsRangeInfo : null;
            var patterns = patternIsRange ? arguments[1].ValueAsRangeInfo : null;
            var replacements = replacementIsRange ? arguments[2].ValueAsRangeInfo : null;

            int textRows = texts != null ? texts.Size.NumberOfRows : 1;
            int textCols = texts != null ? texts.Size.NumberOfCols : 1;
            int patternRows = patterns != null ? patterns.Size.NumberOfRows : 1;
            int patternCols = patterns != null ? patterns.Size.NumberOfCols : 1;
            int replacementsRows = replacements != null ? replacements.Size.NumberOfRows : 1;
            int replacementsCols = replacements != null ? replacements.Size.NumberOfCols : 1;


            var nRows = ExpandedSizeRegexReplace(textRows, patternRows, replacementsRows);
            var nCols = ExpandedSizeRegexReplace(textCols, patternCols, replacementsCols);

            var result = new InMemoryRange(nRows, nCols);

            for (int row = 0; row < nRows; row++)
            {
                for (int col = 0; col < nCols; col++)
                {
                    var textValue = GetRegexReplaceValue(texts, arguments[0], textRows, textCols, row, col);
                    var patternValue = GetRegexReplaceValue(patterns, arguments[1], patternRows, patternCols, row, col);
                    var replacementValue = GetRegexReplaceValue(replacements, arguments[2], replacementsRows, replacementsCols, row, col) ?? string.Empty;

                    if (textValue != null && patternValue == null)
                    {
                        result.SetValue(row, col, ExcelErrorValue.Create(eErrorType.Value));
                    }                        
                    else if (textValue == null || patternValue == null)
                    {
                        result.SetValue(row, col, ExcelErrorValue.Create(eErrorType.NA));
                    }                        
                    else if (caseSensitive > 1 || caseSensitive < 0)
                    {
                        result.SetValue(row, col, ExcelErrorValue.Create(eErrorType.Value));
                    }
                    else
                    {
                        var val = GetRegexReplaced(textValue, patternValue, replacementValue, occurnance, caseSensitive);
                        if (val == null)
                             result.SetValue(row, col, ExcelErrorValue.Create(eErrorType.Value));
                        else
                            result.SetValue(row, col,  val);
                    }                        
                }
            }

            return CreateDynamicArrayResult(result, DataType.ExcelRange);
        }

        private short ExpandedSizeRegexReplace(int a, int b, int c)
        {
            return (short)Math.Max(a, Math.Max(b,c));
        }
        private string GetRegexReplaced(string text, string pattern, string replacement, int occurnance, int caseSensitive)
        {
            if (HasInvalidBackreference(pattern, replacement, (RegexOptions)caseSensitive))
                return null;

            if (Math.Abs(occurnance) > 0)
            {
                var allReplaceMatches = Regex.Matches(text, pattern, (RegexOptions)caseSensitive); 
                var targetIndex = occurnance > 0 ? occurnance - 1 
                                                 : allReplaceMatches.Count + occurnance; // search from end
                if(targetIndex < 0 || targetIndex >= allReplaceMatches.Count)
                {
                    return text;
                }
                var targetMatch = allReplaceMatches[targetIndex];
                return text.Substring(0, targetMatch.Index) 
                    + targetMatch.Result(replacement) 
                    + text.Substring(targetMatch.Index + targetMatch.Length); 
            }
            else
            {
                return Regex.Replace(text, pattern, replacement, (RegexOptions)caseSensitive);
            }            
        }

        private static string GetRegexReplaceValue(
            IRangeInfo range,
            FunctionArgument scalar,
            int argRows, int argCols,
            int row, int col)
        {
            if (range == null)
                return scalar.Value?.ToString() ?? string.Empty;  

            int r = argRows == 1 ? 0 : row;
            int c = argCols == 1 ? 0 : col;
            
            if (r >= argRows || c >= argCols)
                return null;

            return range.GetOffset(r, c)?.ToString() ?? string.Empty;
        }


        private static bool HasInvalidBackreference(string pattern, string replacement, RegexOptions options)
        {
            if (string.IsNullOrEmpty(replacement))
                return false;

            int maxGroup = new Regex(pattern, options).GetGroupNumbers().Max();

            foreach (Match m in Regex.Matches(replacement, @"(?<!\$)\$(\d+)"))
            {
                if (int.TryParse(m.Groups[1].Value, out int refNum) && refNum > maxGroup)
                    return true;
            }
            return false;
        }
    }
}
