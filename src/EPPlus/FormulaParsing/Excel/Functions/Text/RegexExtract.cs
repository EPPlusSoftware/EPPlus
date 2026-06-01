using OfficeOpenXml.FormulaParsing.Excel.Functions.MathFunctions;
using OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup;
using OfficeOpenXml.FormulaParsing.FormulaExpressions;
using OfficeOpenXml.FormulaParsing.Ranges;
using OfficeOpenXml.RichData.IndexRelations;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Text.RegularExpressions;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Text
{
    internal class RegexExtract : RegexFunctionBase
    {
        public override int ArgumentMinLength => 2;

        public override string NamespacePrefix => "_xlfn.";

        public override CompileResult Execute(IList<FunctionArgument> arguments, ParsingContext context)
        {
            bool textIsRange = arguments[0].IsExcelRange;
            bool patternIsRange = arguments[1].IsExcelRange;
            int returnMode = arguments.Count > 2 ? ArgToInt(arguments, 2, 0) : 0;
            int caseSensitivity = arguments.Count > 3 ? ArgToInt(arguments, 3, 0) : 0;

            if (!textIsRange && !patternIsRange)
            {
                var text = arguments[0].Value?.ToString();
                var pattern = arguments[1].Value?.ToString();

                if (text == null || pattern == null)
                    return CreateResult(ExcelErrorValue.Create(eErrorType.NA), DataType.ExcelError);
                if (caseSensitivity > 1 || caseSensitivity < 0 || returnMode < 0 || returnMode > 3)
                    return CreateResult(ExcelErrorValue.Create(eErrorType.Value), DataType.ExcelError);

                if (returnMode == 1)
                {
                    var matches = GetMatches(text, pattern, caseSensitivity);
                    if (matches.Length == 0)
                        return CreateResult(ExcelErrorValue.Create(eErrorType.NA), DataType.ExcelError);

                    var arr = new InMemoryRange((short)1, (short)matches.Length);
                    for (int i = 0; i < matches.Length; i++)
                        arr.SetValue(0, i, matches[i]);

                    return CreateDynamicArrayResult(arr, DataType.ExcelRange);
                }
                else if (returnMode == 2)
                {                    
                    var match = Regex.Match(text, pattern, (RegexOptions)caseSensitivity);
                    if (!match.Success || match.Groups.Count <= 1)
                        return CreateResult(ExcelErrorValue.Create(eErrorType.NA), DataType.ExcelError);

                    var groups = match.Groups
                                      .Cast<Group>()
                                      .Skip(1)
                                      .Select(g => g.Value)
                                      .ToArray();

                    var arr = new InMemoryRange((short)1, (short)groups.Length);
                    for (int i = 0; i < groups.Length; i++)
                        arr.SetValue(0, i, groups[i]);

                    return CreateDynamicArrayResult(arr, DataType.ExcelRange);
                }

                return CreateResult(GetRegexExtractSingle(text, pattern, caseSensitivity), DataType.String);                                    
            }

            // Minst ett range-argument – bygg resultatmatrisen
            var texts = textIsRange ? arguments[0].ValueAsRangeInfo : null;
            var patterns = patternIsRange ? arguments[1].ValueAsRangeInfo : null;

            int textRows = texts != null ? texts.Size.NumberOfRows : 1;
            int textCols = texts != null ? texts.Size.NumberOfCols : 1;
            int patternRows = patterns != null ? patterns.Size.NumberOfRows : 1;
            int patternCols = patterns != null ? patterns.Size.NumberOfCols : 1;

            // Broadcasting-regler:
            //   • Om en dimension är 1  → broadcastas till den andres storlek
            //   • Om båda > 1           → ta max (den kortare ger #N/A vid överflöd)
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
                    else if (Math.Abs(caseSensitivity) > 1 || Math.Abs(returnMode) > 2)
                    {
                        result.SetValue(row, col, ExcelErrorValue.Create(eErrorType.Value));
                    }
                    else
                    {
                        if(returnMode == 2)
                        {
                            var fullMatch = Regex.Match(textValue, patternValue, (RegexOptions)caseSensitivity);
                            var firstMatch = fullMatch.Groups
                                      .Cast<Group>()
                                      .Skip(1)
                                      .Select(g => g.Value)
                                      .ToArray().First().ToString(); // Excel only returns the first match and ignores following matches.
                            result.SetValue(row, col, firstMatch);
                        }
                        else if(returnMode == 1)
                        {
                            var firstMatch = GetMatches(textValue, patternValue, caseSensitivity).First().ToString();
                            result.SetValue(row, col, firstMatch);
                        }
                        else
                        {
                            var match = GetRegexExtractSingle(textValue, patternValue, caseSensitivity);
                            if (match == string.Empty)
                            {
                                result.SetValue(row, col, ExcelErrorValue.Create(eErrorType.NA));
                            }
                            else
                            {
                                result.SetValue(row, col, GetRegexExtractSingle(textValue, patternValue, caseSensitivity));
                            }
                        }                                                
                    }                        
                }
            }

            return CreateDynamicArrayResult(result, DataType.ExcelRange);
        }

        private string[] GetMatches(string text, string pattern, int caseSensitive)
        {
            return Regex.Matches(text, pattern, (RegexOptions)caseSensitive)
                        .Cast<System.Text.RegularExpressions.Match>()
                        .Select(m => m.Value)
                        .ToArray();
        }
        

        private string GetRegexExtractSingle(string text, string pattern, int caseSensitivity)
        {                            
            return Regex.Match(text, pattern, (RegexOptions)caseSensitivity).ToString();
        }

    }
}
