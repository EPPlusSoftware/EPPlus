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
using OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup;
using OfficeOpenXml.FormulaParsing.FormulaExpressions;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Text
{
    [FunctionMetadata(
        Category = ExcelFunctionCategory.Text,
        EPPlusVersion = "4",
        Description = "Returns the length of a supplied text string",
        SupportsArrays = true)]
    internal class Len : ExcelFunction
    {
        public override ExcelFunctionArrayBehaviour ArrayBehaviour => ExcelFunctionArrayBehaviour.FirstArgCouldBeARange;

        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            ValidateArguments(arguments, 1);
            var arg = arguments.First();
            if(arg.Address!=null)
            {
                var currentCell = context.CurrentCell;
                var wsIx = arg.Address.WorksheetIx < 0 ? currentCell.WorksheetIx : arg.Address.WorksheetIx;
                var range = context.ExcelDataProvider.GetRange(arg.Address);
                var firstCell = range.FirstOrDefault();
                if(firstCell != null && firstCell.Value != null)
                {
                    return CreateResult(Convert.ToDouble(firstCell.Value.ToString().Length), DataType.Integer);
                }
                else
                {
                    return CreateResult(0d, DataType.Integer);
                }
            }
            var length = (arg.ValueFirst??"").ToString().Length;
            return CreateResult(Convert.ToDouble(length), DataType.Integer);
        }
    }
}
