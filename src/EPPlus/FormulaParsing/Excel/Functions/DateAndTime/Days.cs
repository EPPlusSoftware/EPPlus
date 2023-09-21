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
using OfficeOpenXml.FormulaParsing.Excel.Functions.Metadata;
using OfficeOpenXml.FormulaParsing.FormulaExpressions;
using System;
using System.Collections.Generic;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.DateAndTime
{
    [FunctionMetadata(
        Category = ExcelFunctionCategory.DateAndTime,
        EPPlusVersion = "4",
        Description = "Calculates the number of days between 2 dates",
        IntroducedInExcelVersion = "2013",
        SupportsArrays = true)]
    internal class Days : ExcelFunction
    {
        public override ExcelFunctionArrayBehaviour ArrayBehaviour => ExcelFunctionArrayBehaviour.Custom;

        public override void ConfigureArrayBehaviour(ArrayBehaviourConfig config)
        {
            config.SetArrayParameterIndexes(0, 1, 2);
        }

        public override int ArgumentMinLength => 2;

        public override string NamespacePrefix => "_xlfn.";

        public override CompileResult Execute(IList<FunctionArgument> arguments, ParsingContext context)
        {
            var numDate1 = ArgToDecimal(arguments, 0, out ExcelErrorValue e1);
            if (e1 != null) return CreateResult(e1.Type);
            var numDate2 = ArgToDecimal(arguments, 1, out ExcelErrorValue e2);
            if (e2 != null) return CreateResult(e2.Type);
            var endDate = DateTime.FromOADate(numDate1);
            var startDate = DateTime.FromOADate(numDate2);
            return CreateResult(endDate.Subtract(startDate).TotalDays, DataType.Date);
        }
    }
}
