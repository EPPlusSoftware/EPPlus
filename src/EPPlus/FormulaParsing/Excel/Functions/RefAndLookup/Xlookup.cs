/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  22/10/2022         EPPlus Software AB           EPPlus v6
 *************************************************************************************************/
using OfficeOpenXml.FormulaParsing.Excel.Functions.Metadata;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup
{
    [FunctionMetadata(
            Category = ExcelFunctionCategory.LookupAndReference,
            EPPlusVersion = "6.0",
            IntroducedInExcelVersion = "2010",
            Description = "Calculates the inverse of the left-tailed probability of the Chi-Square Distribution.")]
    internal class Xlookup : LookupFunction
    {
        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            Stopwatch sw = null;
            if (context.Debug)
            {
                sw = new Stopwatch();
                sw.Start();
            }
            ValidateArguments(arguments, 3);
            var lookupArgs = new LookupArguments(arguments, context);
            var navigator = LookupNavigatorFactory.Create(LookupDirection.Vertical, lookupArgs, context);
            var result = Lookup(navigator, lookupArgs);
            if (context.Debug)
            {
                sw.Stop();
                context.Configuration.Logger.LogFunction("XLOOKUP", sw.ElapsedMilliseconds);
            }
            return result;
        }
    }
}
