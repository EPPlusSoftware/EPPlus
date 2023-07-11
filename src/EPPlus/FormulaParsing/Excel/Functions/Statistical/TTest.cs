/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  06/07/2023         EPPlus Software AB         Implemented function
 *************************************************************************************************/

using OfficeOpenXml.FormulaParsing.Excel.Functions.Helpers;
using OfficeOpenXml.FormulaParsing.Excel.Functions.MathFunctions;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Metadata;
using OfficeOpenXml.FormulaParsing.ExcelUtilities;
using OfficeOpenXml.FormulaParsing.FormulaExpressions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Statistical
{

    [FunctionMetadata(
    Category = ExcelFunctionCategory.Statistical,
    EPPlusVersion = "7.0",
    Description = "Returns the probability for the Student's t-test.")]
    internal class TTest : ExcelFunction
    {
        public override int ArgumentMinLength => 4;

        public override CompileResult Execute(IList<FunctionArgument> arguments, ParsingContext context)
        {

            var array1 = ArgToRangeInfo(arguments, 0);
            var array2 = ArgToRangeInfo(arguments, 1);
            var tails = ArgToDecimal(arguments, 2);
            var type = ArgToDecimal(arguments, 3); //Can be 1, 2 or 3:
                                                   //1. Test is performed paired
                                                   //2. Two-sample equal variance (homoscedastic)
                                                   //3. Two-sample unequal variance (heteroscedastic)

            if ((array1.Size.NumberOfRows * array1.Size.NumberOfCols != array2.Size.NumberOfRows * array2.Size.NumberOfCols)
                && (type == 1))
            {
                return CreateResult(eErrorType.NA); //check if there are equal amount of data points (this is ok if type != 1..?).
            }

            if (tails != 1 && tails != 2)
            {
                return CreateResult(eErrorType.Num); //check if tails input is valid.
            }

            if (type != 1 && type != 2 && type != 3)
            {
                return CreateResult(eErrorType.Num); //check if type input is valid.
            }

            tails = Math.Floor(tails);
            type = Math.Floor(type);
            var tStat = 0d;
            RangeFlattener.GetNumericPairLists(array1, array2, type == 1, out List<double> list1, out List<double> list2);

            if (list1.Count() < 2 || list2.Count() < 2)
            {
                return CreateResult(eErrorType.Div0);
            }

            if (type == 1)
            {
                var differenceList = new List<double>();

                for (var i = 0; i < list1.Count(); i++)
                {
                    differenceList.Add(list1[i] - list2[i]);
                }

                var differenceSD = StandardDeviation(differenceList);
                tStat = differenceList.Average() / (differenceSD / Math.Sqrt(differenceList.Count()));
                tStat = Math.Abs(tStat);

                var result = 1 - StudenttHelper.CumulativeDistributionFunction(tStat, differenceList.Count() - 1);
                return CreateResult(tails == 1 ? result : 2 * result, DataType.Decimal);
            }

            else if (type == 2)
            { 
                var sX = StandardDeviation(list1);
                var sY = StandardDeviation(list2);

                var sXY = Math.Sqrt(((list1.Count() - 1) * Math.Pow(sX, 2) + (list2.Count() - 1) * Math.Pow(sY, 2))
                    / (list1.Count() + list2.Count() - 2));

                tStat = (Math.Abs(list1.Average() - list2.Average())) / (sXY * Math.Sqrt(1d / list1.Count() + 1d / list2.Count()));

                var result = 1 - StudenttHelper.CumulativeDistributionFunction(tStat, list1.Count() + list2.Count() - 2);
                return CreateResult(tails == 1 ? result : 2 * result, DataType.Decimal);

            }

            else
            {
                //Separating the variances instead of combining as in the case of type = 1

                var sX = StandardDeviation(list1);
                var sY = StandardDeviation(list2);

                var varX = Math.Pow(sX, 2);
                var varY = Math.Pow(sY, 2);

                var meanX = list1.Average();
                var meanY = list2.Average();

                //For type = 3 (Welsh-test), the degrees of freedom is calculated differently.

                var numerator = Math.Pow(varX / list1.Count() + varY / list2.Count(), 2);

                var denominator = Math.Pow(varX / list1.Count(), 2) / (list1.Count() - 1) + Math.Pow(varY / list2.Count(), 2) / (list2.Count() - 1);

                var degreesOfFreedom = numerator/ denominator; //We do not truncate here for excel compliance.

                tStat = (meanX - meanY) / Math.Sqrt(varX / list1.Count() + varY/ list2.Count());

                tStat = Math.Abs(tStat);

                var result = 1 - StudenttHelper.CumulativeDistributionFunction(tStat, degreesOfFreedom);
                return CreateResult(tails == 1 ? result : result * 2, DataType.Decimal);

            }
        }

        internal static double StandardDeviation(List<double> values)
        {
            //Returns the standard deviation of a list

            var std = 0d;
            var mean = values.Average();

            for (var i = 0; i < values.Count; i++)
            {
                std += Math.Pow(values[i] - mean, 2);
            }

            std = Math.Sqrt(std / (values.Count() - 1));

            return std;
        }
    }
}
