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
using OfficeOpenXml.FormulaParsing.Utilities;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.MathFunctions
{
    [FunctionMetadata(
        Category = ExcelFunctionCategory.Statistical,
        EPPlusVersion = "4",
        Description = "Returns the smallest value from a list of supplied values, counting text and the logical value FALSE as the value 0 and counting the logical value TRUE as the value 1")]
    internal class Mina : ExcelFunction
    {
        private readonly DoubleEnumerableArgConverter _argConverter;

        public Mina()
            : this(new DoubleEnumerableArgConverter())
        {

        }
        public Mina(DoubleEnumerableArgConverter argConverter)
        {
            Require.That(argConverter).Named("argConverter").IsNotNull();
            _argConverter = argConverter;
        }
        public override int ArgumentMinLength => 1;
        public override CompileResult Execute(IList<FunctionArgument> arguments, ParsingContext context)
        {
            var values = _argConverter.ConvertArgsIncludingOtherTypes(arguments, false);
            return CreateResult(values.Min(), DataType.Decimal);
        }
    }
}
