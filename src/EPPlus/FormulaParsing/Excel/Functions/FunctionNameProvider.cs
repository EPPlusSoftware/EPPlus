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

namespace OfficeOpenXml.FormulaParsing.Excel.Functions
{
    /// <summary>
    /// Function name provider
    /// </summary>
    public class FunctionNameProvider : IFunctionNameProvider
    {
        private FunctionNameProvider()
        {

        }
        /// <summary>
        /// Empty
        /// </summary>
        public static FunctionNameProvider Empty
        {
            get { return new FunctionNameProvider(); }
        }
        /// <summary>
        /// Is function name
        /// </summary>
        /// <param name="name"></param>
        /// <returns></returns>
        public virtual bool IsFunctionName(string name)
        {
            return false;
        }
    }
}
