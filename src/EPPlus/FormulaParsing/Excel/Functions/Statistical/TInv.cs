using OfficeOpenXml.FormulaParsing.Excel.Functions.Helpers;
using OfficeOpenXml.FormulaParsing.FormulaExpressions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Statistical
{
    internal class TInv : ExcelFunction
    {
        public override int ArgumentMinLength => 2;

        public override CompileResult Execute(IList<FunctionArgument> arguments, ParsingContext context)
        {

            var probability = ArgToDecimal(arguments, 0);
            var degreesOfFreedom = ArgToDecimal(arguments, 1);

            degreesOfFreedom = System.Math.Floor(degreesOfFreedom);

            if (probability <= 0 || probability > 1)
            {
                return CreateResult(eErrorType.Num);
            }

            if (degreesOfFreedom < 1)
            {
                return CreateResult(eErrorType.Num);
            }

            return CreateResult(StudenttHelper.InverseTFunc(probability, degreesOfFreedom), DataType.Decimal);
        }
    }
}
