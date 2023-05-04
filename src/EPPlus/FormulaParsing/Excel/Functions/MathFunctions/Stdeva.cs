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
using OfficeOpenXml.FormulaParsing.FormulaExpressions;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Metadata;
using OfficeOpenXml.Utils;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.MathFunctions
{
    [FunctionMetadata(
        Category = ExcelFunctionCategory.Statistical,
        EPPlusVersion = "5.5",
        Description = "Returns the standard deviation of a supplied set of values (which represent a sample of a population), counting text and the logical value FALSE as the value 0 and counting the logical value TRUE as the value 1")]
    internal class Stdeva : HiddenValuesHandlingFunction
    {
        private readonly DoubleEnumerableArgConverter _argConverter;

        public Stdeva()
            : this(new DoubleEnumerableArgConverter())
        {

        }

        public Stdeva(DoubleEnumerableArgConverter argConverter)
        {
            Require.Argument(argConverter).IsNotNull("argConverter");
            _argConverter = argConverter;
        }

        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            ValidateArguments(arguments, 1);
            var values = _argConverter.ConvertArgsIncludingOtherTypes(arguments, IgnoreHiddenValues).Select(x => (double)x);
            return StandardDeviation(values);
        }

        private CompileResult StandardDeviation(IEnumerable<double> values)
        {
            double ret = 0;
            if (values.Any())
            {
                var nValues = values.Count();
                if (nValues == 1) throw new ExcelErrorValueException(eErrorType.Div0);
                //Compute the Average       
                double avg = values.Average();
                //Perform the Sum of (value-avg)_2_2       
                double sum = values.Sum(d => MathObj.Pow(d - avg, 2));
                //Put it all together       
                var div = Divide(sum, (values.Count() - 1));
                if (div is ExcelErrorValue e) return CreateResult(e, DataType.ExcelError);

                ret = MathObj.Sqrt((double)div);
            }
            return CreateResult(ret, DataType.Decimal);
        }

    }
}
