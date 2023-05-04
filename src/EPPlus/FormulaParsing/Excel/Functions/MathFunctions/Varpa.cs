/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  05/25/2020         EPPlus Software AB       Implemented function
 *************************************************************************************************/
using OfficeOpenXml.FormulaParsing.Excel.Functions.Metadata;
using OfficeOpenXml.FormulaParsing.FormulaExpressions;
using OfficeOpenXml.Utils;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.MathFunctions
{
    [FunctionMetadata(
         Category = ExcelFunctionCategory.Statistical,
         EPPlusVersion = "5.5",
         Description = "Returns the variance of a supplied set of values (which represent a sample of a population), counting text and the logical value FALSE as the value 0 and counting the logical value TRUE as the value 1")]
    internal class Varpa : ExcelFunction
    {
        private readonly DoubleEnumerableArgConverter _argConverter;

        public Varpa()
            : this(new DoubleEnumerableArgConverter())
        {

        }

        public Varpa(DoubleEnumerableArgConverter argConverter)
        {
            Require.Argument(argConverter).IsNotNull("argConverter");
            _argConverter = argConverter;
        }

        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            if (!arguments.Any() || arguments.Count() < 2) return CreateResult(eErrorType.Div0);
            var values = _argConverter.ConvertArgsIncludingOtherTypes(arguments, false);
            var result = VarMethods.VarP(values);
            return CreateResult(result, DataType.Decimal);
        }
    }
}
