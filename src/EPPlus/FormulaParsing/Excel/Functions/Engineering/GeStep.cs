
using OfficeOpenXml.FormulaParsing.Excel.Functions.Metadata;
using OfficeOpenXml.FormulaParsing.FormulaExpressions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Engineering
{
    [FunctionMetadata(
       Category = ExcelFunctionCategory.Engineering,
       EPPlusVersion = "7.0",
       Description = "Returns 1 if number ≥ step; returns 0 (zero) otherwise. Use this function to filter a set of values. For example, by summing several GESTEP functions you calculate the count of values that exceed a threshold")]
    internal class GeStep : ExcelFunction
    {
        public override int ArgumentMinLength => 1;

        public override CompileResult Execute(IList<FunctionArgument> arguments, ParsingContext context)
        {
            var arg1 = ArgToDecimal(arguments, 0);
            var arg2 = 0d;
            if (arguments.Count > 1)
            {
                arg2 = ArgToDecimal(arguments, 1);
            }
            
            if (arg1 >= arg2)
            {
               return CreateResult(1, DataType.Integer);
            }
            else
            {
               return CreateResult(0, DataType.Integer);
            }
        }
    }
}
