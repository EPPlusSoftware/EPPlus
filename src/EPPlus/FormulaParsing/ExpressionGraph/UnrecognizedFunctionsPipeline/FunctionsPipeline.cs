/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  06/15/2020         EPPlus Software AB       EPPlus 5.2
 *************************************************************************************************/
using OfficeOpenXml.FormulaParsing.Excel.Functions;
using OfficeOpenXml.FormulaParsing.ExpressionGraph.UnrecognizedFunctionsPipeline.Handlers;
using System;
using System.Collections.Generic;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.ExpressionGraph.UnrecognizedFunctionsPipeline
{
    /// <summary>
    /// A pipeline where handlers for unrecognized function names are registred.
    /// </summary>
    internal class FunctionsPipeline
    {
        public FunctionsPipeline()
            : this(new RangeOffsetFunctionHandler())
        {

        }

        public FunctionsPipeline(params UnrecognizedFunctionsHandler[] handlers)
        {
            _handlers = handlers;
        }

        private IEnumerable<UnrecognizedFunctionsHandler> _handlers;

        /// <summary>
        /// Tries to find a registred handler that can handle the function name
        /// If success this <see cref="ExcelFunction"/> are returned.
        /// </summary>
        /// <param name="funcName">The unrecognized function name</param>
        /// <returns>An <see cref="ExcelFunction"/> that can handle the function call</returns>
        internal ExcelFunction FindFunction(string funcName)
        {
            foreach(var handler in _handlers)
            {
                if(handler.Handle(funcName, out ExcelFunction function))
                {
                    return function;
                }
            }
            return default;
        }

    }
}
