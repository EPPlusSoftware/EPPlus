using OfficeOpenXml.FormulaParsing.FormulaExpressions;
using OfficeOpenXml.FormulaParsing.Ranges;
using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Text
{
    internal class RegexTest : ExcelFunction
    {
        public override int ArgumentMinLength => 2;
        public override string NamespacePrefix => "_xlfn.";
        
        public override CompileResult Execute(IList<FunctionArgument> arguments, ParsingContext context)
        {
            bool textIsRange = arguments[0].IsExcelRange;
            bool patternIsRange = arguments[1].IsExcelRange;
            int caseSensitivity = ArgToInt(arguments, 2, 0);

            if (!textIsRange && !patternIsRange)
            {
                // Skalär × skalär – ursprungligt beteende
                var text = arguments[0].Value?.ToString();
                var pattern = arguments[1].Value?.ToString();

                if (text == null || pattern == null)
                    return CreateResult(ExcelErrorValue.Create(eErrorType.NA), DataType.ExcelError);

                return CreateResult(GetRegexTest(text, pattern, caseSensitivity), DataType.Boolean);
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
                    else if(Math.Abs(caseSensitivity) <
                    else
                        result.SetValue(row, col, GetRegexTest(textValue, patternValue));
                }
            }

            return CreateDynamicArrayResult(result, DataType.ExcelRange);
        }

        /// <summary>
        /// Hämtar strängvärdet för (row, col) med broadcasting.
        /// Returnerar null om cellen är utanför räckvidden (→ #N/A).
        /// </summary>
        private static string GetValue(
            IRangeInfo range,
            FunctionArgument scalar,
            int nRows, int nCols,
            int row, int col)
        {
            if (range == null)
                // Skalärargument – broadcastas alltid
                return scalar.Value?.ToString();

            // Beräkna verkligt index med broadcasting (storlek 1 → använd index 0)
            int r = nRows == 1 ? 0 : row;
            int c = nCols == 1 ? 0 : col;

            // Utanför räckvidden → #N/A
            if (r >= nRows || c >= nCols)
                return null;

            return range.GetOffset(r, c)?.ToString();
        }

        /// <summary>
        /// Beräknar resultatdimensionen för en axel enligt Excels broadcasting-regler.
        /// </summary>
        private static short ExpandedSize(int a, int b)
        {
            if (a == 1) return (short)b;
            if (b == 1) return (short)a;
            return (short)Math.Max(a, b);   // Båda > 1: max-storlek, överskott → #N/A
        }

        private static bool GetRegexTest(string text, string pattern, int caseSensitive)
            => Regex.IsMatch(text, pattern);
    }
}