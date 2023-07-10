using OfficeOpenXml.FormulaParsing.FormulaExpressions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Statistical
{
    internal class TTest : ExcelFunction
    {
        public override int ArgumentMinLength => 4;

        public override CompileResult Execute(IList<FunctionArgument> arguments, ParsingContext context)
        {

            var array1 = ArgToRangeInfo(arguments, 0);
            var array2 = ArgToRangeInfo(arguments, 1);
            var tails = ArgToDecimal(arguments, 2);
            var type = ArgToDecimal(arguments, 3);

            if (array1.Size.NumberOfRows * array1.Size.NumberOfCols != array2.Size.NumberOfRows * array2.Size.NumberOfCols)
            {
                return CreateResult(eErrorType.Num); //check if num is the correct error type.
            }

            List<double>list1 = new List<double>();
            List<double>list2 = new List<double>();

            for (var index = 0; index <= )

            //Returns probability associated with a Students t-test.
            //Uses data in Array 1 & 2 to compute non-negative t-statistics. If tails = 1 T.TEST returns the probability
            //of a higher value of the t-statistics under the assumption that array1 and array2 are samples from populations
            //with the same mean. The value returned when tails = 2 is double that returned when tails = 1 and corresponds
            //to the probability of a higher absolute value of the t-statistics under "the same population means" assumption.
            throw new NotImplementedException();
        }
    }
}
