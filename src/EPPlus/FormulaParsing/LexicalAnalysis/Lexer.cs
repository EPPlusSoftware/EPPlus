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
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml.FormulaParsing.Excel.Functions;
//using OfficeOpenXml.FormulaParsing.Excel.Functions;

namespace OfficeOpenXml.FormulaParsing.LexicalAnalysis
{
    public class Lexer : ILexer
    {
        public Lexer(FunctionRepository functionRepository, INameValueProvider nameValueProvider)
            :this(new SourceCodeTokenizer(functionRepository, nameValueProvider), new SyntacticAnalyzer())
        {

        }

        public Lexer(ISourceCodeTokenizer tokenizer, ISyntacticAnalyzer analyzer)
        {
            _tokenizer = tokenizer;
            _analyzer = analyzer;
        }

        private readonly ISourceCodeTokenizer _tokenizer;
        private readonly ISyntacticAnalyzer _analyzer;
        public IList<Token> Tokenize(string input)
        {
            return Tokenize(input, null);
        }
        public IList<Token> Tokenize(string input, string worksheet)
        {
            var tokens = _tokenizer.Tokenize(input, worksheet);
            _analyzer.Analyze(tokens);
            return tokens;
        }
    }
}
