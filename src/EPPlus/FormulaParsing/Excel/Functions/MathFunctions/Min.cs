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
using OfficeOpenXml.FormulaParsing.FormulaExpressions;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.MathFunctions
{
    [FunctionMetadata(
        Category = ExcelFunctionCategory.Statistical,
        EPPlusVersion = "4",
        Description = "Returns the smallest value from a list of supplied numbers")]
    internal class Min : HiddenValuesHandlingFunction
    {
        public Min() : base()
        {
            IgnoreErrors = false;
        }
        public override ExcelFunctionParametersInfo ParametersInfo => new ExcelFunctionParametersInfo(new Func<int, FunctionParameterInformation>((argumentIndex) =>
        {
            return FunctionParameterInformation.IgnoreErrorInPreExecute;
        }));

        public override int ArgumentMinLength => 1;
        public override CompileResult Execute(IList<FunctionArgument> arguments, ParsingContext context)
        {
            var values = ArgsToDoubleEnumerable(arguments, context, x =>
            {
                x.IgnoreHiddenCells = IgnoreHiddenValues;
                x.IgnoreErrors = IgnoreErrors;
                x.IgnoreNestedSubtotalAggregate = IgnoreNestedSubtotalsAndAggregates;
                x.IgnoreNonNumeric = true;
            }, out ExcelErrorValue e1);
            if (e1 != null) return CompileResult.GetErrorResult(e1.Type);
            if(!values.Any())
            {
                return CreateResult(0d, DataType.Decimal);
            }
            return CreateResult(values.Min(), DataType.Decimal);
        }
    }
}
