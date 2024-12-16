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
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.FormulaExpressions.CompileResults
{
    /// <summary>
    /// Local image compile result
    /// </summary>
    public class LocalImageCompileResult : AddressCompileResult
    {
        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="result"></param>
        /// <param name="dataType"></param>
        /// <param name="address"></param>
        public LocalImageCompileResult(object result,  FormulaRangeAddress address) : base(result, DataType.LocalImage, address)
        {
        }

        /// <summary>
        /// The result is a local image
        /// </summary>
        public override CompileResultType ResultType
        {
            get
            {
                return CompileResultType.LocalImage;
            }
        }
    }
}
