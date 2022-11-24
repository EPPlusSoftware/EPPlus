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
using OfficeOpenXml.FormulaParsing.Excel.Functions.Metadata;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using static OfficeOpenXml.FormulaParsing.ExcelDataProvider;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup
{
    [FunctionMetadata(
        Category = ExcelFunctionCategory.LookupAndReference,
        EPPlusVersion = "4",
        Description = "Returns one of a list of values, depending on the value of a supplied index number")]
    internal class Choose : ExcelFunction
    {
        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            ValidateArguments(arguments, 2);
            var items = new List<object>();
            for (int x = 0; x < arguments.Count(); x++)
            {
                items.Add(arguments.ElementAt(x).ValueFirst);
            }

            var chooseIndices = arguments.ElementAt(0).ValueFirst as IEnumerable<FunctionArgument>;
            if (chooseIndices != null && chooseIndices.Count() > 1)
            {
                IntArgumentParser intParser = new IntArgumentParser();
                object[] values = chooseIndices.Select(chosenIndex => items[(int)intParser.Parse(chosenIndex.ValueFirst)]).ToArray();
                return CreateResult(values, DataType.Enumerable);
            }
            else
            {
                var index = ArgToInt(arguments, 0);
                var choosedValue = arguments.ElementAt(index).Value;
                if(choosedValue is IRangeInfo)
                {
                    return CreateResult(choosedValue, DataType.Enumerable);
                }
                return CompileResultFactory.Create(choosedValue);
            }
        }
        public override bool ReturnsReference => true;
    }
}
