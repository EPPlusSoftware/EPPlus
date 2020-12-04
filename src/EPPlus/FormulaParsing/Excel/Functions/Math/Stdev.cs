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
using OfficeOpenXml.FormulaParsing.Exceptions;
using MathObj = System.Math;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Metadata;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Math
{
    [FunctionMetadata(
        Category = ExcelFunctionCategory.Statistical,
        EPPlusVersion = "4",
        Description = "Returns the standard deviation of a supplied set of values (which represent a sample of a population) ")]
    internal class Stdev : HiddenValuesHandlingFunction
    {
        public Stdev() : base()
        {
            IgnoreErrors = false;
        }
        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            ValidateArguments(arguments, 1);
            var values = ArgsToDoubleEnumerable(arguments, context, IgnoreErrors).Select(x => (double)x);
            return CreateResult(StandardDeviation(values), DataType.Decimal);
        }

        private double StandardDeviation(IEnumerable<double> values)
        {
            double ret = 0;
            if (values.Any())
            {
                var nValues = values.Count();
                if(nValues == 1) throw new ExcelErrorValueException(eErrorType.Div0);
                //Compute the Average       
                double avg = values.Average();
                //Perform the Sum of (value-avg)_2_2       
                double sum = values.Sum(d => MathObj.Pow(d - avg, 2));
                //Put it all together       
                ret = MathObj.Sqrt(Divide(sum, (values.Count() - 1)));
            }
            return ret;
        } 

    }
}
