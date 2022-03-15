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
            IntroducedInExcelVersion = "2016",
            Description = "Searches a range or an array, and then returns the item corresponding to the first match it finds. Will return a VALUE error if the functions returns an array (EPPlus does not support dynamic arrayformulas)")]
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
            var lookupValue = arguments.ElementAt(0).Value;
            var lookupArray = Enumerable.Empty<object>().ToArray();
            if(arguments.ElementAt(1).IsExcelRange)
            {
                lookupArray = arguments.ElementAt(1).ValueAsRangeInfo.Select(x => x.Value).ToArray();
            }
            else
            {
                lookupArray = ArgsToObjectEnumerable(true, new List<FunctionArgument> { arguments.ElementAt(1) }, context).ToArray();
            }
            var returnArray = Enumerable.Empty<object[]>();
            if (arguments.ElementAt(1).IsExcelRange)
            {
                lookupArray = arguments.ElementAt(2).ValueAsRangeInfo.Select(x => x.Value).ToArray();
            }
            else
            {
                lookupArray = ArgsToObjectEnumerable(true, new List<FunctionArgument> { arguments.ElementAt(2) }, context).ToArray();
            }
            if (context.Debug)
            {
                sw.Stop();
                context.Configuration.Logger.LogFunction("XLOOKUP", sw.ElapsedMilliseconds);
            }
            return null;
        }
    }
}
