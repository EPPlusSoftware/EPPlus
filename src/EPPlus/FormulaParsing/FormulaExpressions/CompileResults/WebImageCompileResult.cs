/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  10/29/2024         EPPlus Software AB       Initial release EPPlus 8
 *************************************************************************************************/
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;

namespace OfficeOpenXml.FormulaParsing.FormulaExpressions.CompileResults
{
    /// <summary>
    /// Local image compile result
    /// </summary>
    public class WebImageCompileResult : CompileResult
    {
        /// <summary>
        /// Web image compile result
        /// </summary>
        public WebImageCompileResult(object result) 
            : base(result, DataType.WebImage)
        {
            
        }

        /// <summary>
        /// Web image compile result
        /// </summary>
        public WebImageCompileResult(object result, ParsingContext ctx)
            : base(result, DataType.WebImage)
        {
            
        }

        /// <summary>
        /// The result is a web image
        /// </summary>
        public override CompileResultType ResultType
        {
            get
            {
                return CompileResultType.WebImage;
            }
        }
    }
}
