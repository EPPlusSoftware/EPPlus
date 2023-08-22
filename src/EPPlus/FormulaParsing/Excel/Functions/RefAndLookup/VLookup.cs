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
using System.Diagnostics;
using System.Linq;
using System.Text;
using OfficeOpenXml.FormulaParsing.FormulaExpressions;
using OfficeOpenXml.FormulaParsing.ExcelUtilities;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Metadata;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup
{
    [FunctionMetadata(
        Category = ExcelFunctionCategory.LookupAndReference,
        EPPlusVersion = "4",
        Description = "Looks up a supplied value in the first column of a table, and returns the corresponding value from another column")]
    internal class VLookup : LookupFunction
    {
        public override FunctionParameterInformation GetParameterInfo(int argumentIndex)
        {
            if(argumentIndex == 1)
            {
                return FunctionParameterInformation.IgnoreAddress; //
            }
            return FunctionParameterInformation.Normal;
        }
        public override int ArgumentMinLength => 3;
        public override CompileResult Execute(IList<FunctionArgument> arguments, ParsingContext context)
        {
            Stopwatch sw = null;
            if (context.Debug)
            {
                sw = new Stopwatch();
                sw.Start();
            }
            var arg1 = arguments[1];
            if (arg1.DataType == DataType.ExcelError) return CompileResult.GetErrorResult(((ExcelErrorValue)arg1.Value).Type);
            var lookupArgs = new LookupArguments(arguments, context);
            var navigator = LookupNavigatorFactory.Create(LookupDirection.Vertical, lookupArgs, context);
            var result = Lookup(navigator, lookupArgs);
            if (context.Debug)
            {
                sw.Stop();
                context.Configuration.Logger.LogFunction("VLOOKUP", sw.ElapsedMilliseconds);
            }
            return result;
        }
    }
}
