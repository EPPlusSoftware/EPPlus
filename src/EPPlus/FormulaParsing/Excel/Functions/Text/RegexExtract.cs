using OfficeOpenXml.FormulaParsing.Excel.Functions.MathFunctions;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Metadata;
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
    [FunctionMetadata(
        Category = ExcelFunctionCategory.Text,
        EPPlusVersion = "8.6",
        Description = "Extracts text matching a regular expression pattern from input string values.",
        SupportsArrays = true)]
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
                if (caseSensitivity > 1 || caseSensitivity < 0 || returnMode < 0 || returnMode > 2)
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
                    // Read the number of capturing groups from the pattern (GetGroupNumbers
                    // includes group 0). A failed match reports Groups.Count == 1, so this must
                    // not be read from the match. No groups -> #VALUE!; groups but no match -> #N/A.
                    var regex = new Regex(pattern, (RegexOptions)caseSensitivity);
                    if (regex.GetGroupNumbers().Length <= 1)
                        return CreateResult(ExcelErrorValue.Create(eErrorType.Value), DataType.ExcelError);

                    var match = regex.Match(text);
                    if (!match.Success)
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
                var firstMatch = Regex.Match(text, pattern, (RegexOptions)caseSensitivity);
                if (!firstMatch.Success)
                    return CreateResult(ExcelErrorValue.Create(eErrorType.NA), DataType.ExcelError);
                return CreateResult(firstMatch.Value, DataType.String);
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
                    {
                        result.SetValue(row, col, ExcelErrorValue.Create(eErrorType.NA));
                    }
                    // Use the same validation as the scalar branch. The previous Math.Abs check
                    // let negative arguments through, which fell into mode 0 or reached
                    // (RegexOptions)(-1). Excel returns #VALUE! per cell for these.
                    else if (caseSensitivity > 1 || caseSensitivity < 0 || returnMode < 0 || returnMode > 2)
                    {
                        result.SetValue(row, col, ExcelErrorValue.Create(eErrorType.Value));
                    }
                    else
                    {
                        // Compute per cell and catch invalid-pattern exceptions here so that a
                        // single bad cell becomes #VALUE! in place, while the other cells are
                        // still calculated (verified against Excel).
                        try
                        {
                            var options = (RegexOptions)caseSensitivity;
                            if (returnMode == 2)
                            {
                                // A failed match reports Groups.Count == 1, so the number of
                                // capturing groups must be read from the pattern itself (via
                                // GetGroupNumbers, which includes group 0) rather than from the
                                // match. No groups -> #VALUE!; groups but no match -> #N/A.
                                var regex = new Regex(patternValue, options);
                                if (regex.GetGroupNumbers().Length <= 1)
                                {
                                    result.SetValue(row, col, ExcelErrorValue.Create(eErrorType.Value));
                                }
                                else
                                {
                                    var match = regex.Match(textValue);
                                    if (!match.Success)
                                    {
                                        result.SetValue(row, col, ExcelErrorValue.Create(eErrorType.NA));
                                    }
                                    else
                                    {
                                        // In range mode only the first group is returned per cell.
                                        result.SetValue(row, col, match.Groups[1].Value);
                                    }
                                }
                            }
                            else if (returnMode == 1)
                            {
                                var matches = GetMatches(textValue, patternValue, caseSensitivity);
                                if (matches.Length == 0)
                                {
                                    result.SetValue(row, col, ExcelErrorValue.Create(eErrorType.NA));
                                }
                                else
                                {
                                    // In range mode only the first match is returned per cell.
                                    result.SetValue(row, col, matches[0]);
                                }
                            }
                            else
                            {
                                var match = Regex.Match(textValue, patternValue, options);
                                if (!match.Success)
                                {
                                    result.SetValue(row, col, ExcelErrorValue.Create(eErrorType.NA));
                                }
                                else
                                {
                                    result.SetValue(row, col, match.Value);
                                }
                            }
                        }
                        catch (ArgumentException)
                        {
                            // Invalid regex pattern in this cell -> #VALUE! for this cell only.
                            result.SetValue(row, col, ExcelErrorValue.Create(eErrorType.Value));
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
    }
}