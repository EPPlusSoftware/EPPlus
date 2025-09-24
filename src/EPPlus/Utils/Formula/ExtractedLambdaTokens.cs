using OfficeOpenXml.FormulaParsing.FormulaExpressions;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Utils.Formula
{
    internal class ExtractedLambdaTokens
    {
        public RpnTokens LambdaTokens { get; set; }

        public List<Token> VariableTokens { get; set; }

        public string LambdaFormula { get; set; }
    }
}
