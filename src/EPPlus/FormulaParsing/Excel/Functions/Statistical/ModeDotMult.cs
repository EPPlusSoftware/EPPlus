/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  06/23/2020         EPPlus Software AB       EPPlus 5.2
 *************************************************************************************************/
using OfficeOpenXml.FormulaParsing.Excel.Functions.Metadata;
using OfficeOpenXml.FormulaParsing.FormulaExpressions;
using OfficeOpenXml.FormulaParsing.Ranges;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Statistical
{
    [FunctionMetadata(
   Category = ExcelFunctionCategory.Statistical,
   EPPlusVersion = "7",
   Description = "Returns a vertical array of the most frequently occurring, or repetitive values in an array or range of data.")]
    internal class ModeDotMult : HiddenValuesHandlingFunction
    {

        private class ModeValue
        {
            public ModeValue(double variable, int quantity, int sortOrder)
            {
                Quantity = quantity;
                SortOrder = sortOrder;
                Variable = variable;
            }
            public int Quantity{ get; set; }

            public int SortOrder { get; set; }

            public double Variable { get; set; }
            public void Increase()
            {
                Quantity++;
            }
        }

        public ModeDotMult() 
        {
            IgnoreErrors = false;
        }
        /// <summary>
        /// Reference Parameters do not need to be follows in the dependency chain.
        /// </summary>
        public override ExcelFunctionParametersInfo ParametersInfo => new ExcelFunctionParametersInfo(new Func<int, FunctionParameterInformation>((argumentIndex) =>
        {
            return FunctionParameterInformation.IgnoreErrorInPreExecute;
        }));
        public override string NamespacePrefix => "_xlfn.";
        public override int ArgumentMinLength => 1;

        public override CompileResult Execute(IList<FunctionArgument> arguments, ParsingContext context)
        {
            if (arguments.Count() > 255)
            {
                return CompileResult.GetErrorResult(eErrorType.Value);
            }
            var numbers = ArgsToDoubleEnumerable(IgnoreHiddenValues, IgnoreErrors, arguments, context);
            var countNumbers = new Dictionary<double, ModeValue>();
            int maxCount = 0;
            var sortOrder = 1;
            foreach (var variable in numbers)
            {
                if (!countNumbers.ContainsKey(variable))
                {
                    countNumbers[variable] = new ModeValue(variable, 1, sortOrder);
                    sortOrder++;
                }
                else
                {
                    countNumbers[variable].Increase();
                }
                if (countNumbers[variable].Quantity > maxCount)
                {
                    maxCount = countNumbers[variable].Quantity;
                }
            }
            if (maxCount == 0)
            {
                return CreateResult(eErrorType.Num);
            }
            var maxNumbers = countNumbers.Values.Where(x => x.Quantity == maxCount).ToList();
            maxNumbers.Sort((a, b) => a.SortOrder.CompareTo(b.SortOrder));
            var result = new InMemoryRange(maxNumbers.Count, 1);
            for (var row = 0; row < maxNumbers.Count; row++)
            {
                var number = maxNumbers[row];
                result.SetValue(row, 0, number.Variable);
            }
            return CreateResult(result, DataType.ExcelRange);
            
        }
    }
}
