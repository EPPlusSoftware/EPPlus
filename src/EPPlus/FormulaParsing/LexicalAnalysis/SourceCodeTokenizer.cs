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
using System.Globalization;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using OfficeOpenXml.FormulaParsing.Excel.Functions;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis.PostProcessing;

namespace OfficeOpenXml.FormulaParsing.LexicalAnalysis
{
    public class SourceCodeTokenizer : ISourceCodeTokenizer
    {
        public static ISourceCodeTokenizer Default
        {
            get { return new SourceCodeTokenizer(FunctionNameProvider.Empty, NameValueProvider.Empty, false); }
        }
        public static ISourceCodeTokenizer R1C1
        {
            get { return new SourceCodeTokenizer(FunctionNameProvider.Empty, NameValueProvider.Empty, true); }
        }

        public static ISourceCodeTokenizer Optimized
        {
            get { return new OptimizedSourceCodeTokenizer(FunctionNameProvider.Empty, NameValueProvider.Empty, false); }
        }

        public SourceCodeTokenizer(IFunctionNameProvider functionRepository, INameValueProvider nameValueProvider, bool r1c1 = false)
            : this(new TokenFactory(functionRepository, nameValueProvider, r1c1))
        {
            _nameValueProvider = nameValueProvider;
        }
        public SourceCodeTokenizer(ITokenFactory tokenFactory)
        {
            _tokenFactory = tokenFactory;
        }

        private readonly ITokenFactory _tokenFactory;
        private readonly INameValueProvider _nameValueProvider;

        public IList<Token> Tokenize(string input)
        {
            return Tokenize(input, null);
        }
        public IList<Token> Tokenize(string input, string worksheet)
        {
            if (string.IsNullOrEmpty(input))
            {
                return new List<Token>();
            }
            // MA 1401: Ignore leading plus in formula.
            input = input.TrimStart('+');
            var context = new TokenizerContext(input, worksheet, _tokenFactory);
            var handler = context.CreateHandler(_nameValueProvider);
            while (handler.HasMore())
            {
                handler.Next();
            }
            context.PostProcess();

            return context.Result;
        }
    }
}
