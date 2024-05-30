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

namespace OfficeOpenXml.FormulaParsing.LexicalAnalysis
{
    /// <summary>
    /// Source code tokenizer
    /// </summary>
    public interface ISourceCodeTokenizer
    {
        /// <summary>
        /// Tokenize
        /// </summary>
        /// <param name="input"></param>
        /// <param name="worksheet"></param>
        /// <returns></returns>
        IList<Token> Tokenize(string input, string worksheet);
        /// <summary>
        /// Tokenize
        /// </summary>
        /// <param name="input"></param>
        /// <returns></returns>
        IList<Token> Tokenize(string input);
    }
}
