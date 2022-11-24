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
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup;
using OfficeOpenXml.FormulaParsing.ExpressionGraph.FunctionCompilers;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.ExpressionGraph.UnrecognizedFunctionsPipeline.Handlers
{
    /// <summary>
    /// Handles a range, where the second argument is a call to the OFFSET function
    /// Example: A1:OFFSET(B2, 2, 0).
    /// </summary>
    internal class RangeOffsetFunctionHandler : UnrecognizedFunctionsHandler
    {
        public override bool Handle(string funcName, IEnumerable<Expression> children, ParsingContext context, out ExcelFunction function)
        {
            function = null;
            if(funcName.Contains(":OFFSET"))
            {
                var functionCompilerFactory = new FunctionCompilerFactory(context.Configuration.FunctionRepository, context);
                var startRange = funcName.Split(':')[0];
                var c = context.Scopes.Current;
                var resultRange = context.ExcelDataProvider.GetRange(c.Address.WorksheetName, c.Address.FromRow, c.Address.FromCol, startRange);
                var rangeOffset = new RangeOffset
                {
                    StartRange = resultRange
                };
                var compiler = functionCompilerFactory.Create(new Offset());
                children.First().Children.First().IgnoreCircularReference = true;
                var compileResult = compiler.Compile(children);
                rangeOffset.EndRange = compileResult.Result as IRangeInfo;
                function = rangeOffset;
                return true;
            }
            return false;
        }
    }
}
