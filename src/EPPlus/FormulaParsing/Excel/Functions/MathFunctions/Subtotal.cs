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

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.MathFunctions
{
    [FunctionMetadata(
        Category = ExcelFunctionCategory.MathAndTrig,
        EPPlusVersion = "4",
        Description = "Performs a specified calculation (e.g. the sum, product, average, etc.) for a supplied set of values")]
    internal class Subtotal : ExcelFunction
    {
        private Dictionary<int, HiddenValuesHandlingFunction> _functions = new Dictionary<int, HiddenValuesHandlingFunction>();
        
        public Subtotal()
        {
            Initialize();
        }

        private void Initialize()
        {
            _functions[1] = new Average();
            _functions[2] = new Count();
            _functions[3] = new CountA();
            _functions[4] = new Max();
            _functions[5] = new Min();
            _functions[6] = new Product();
            _functions[7] = new Stdev();
            _functions[8] = new StdevP();
            _functions[9] = new SumSubtotal();
            _functions[10] = new Var();
            _functions[11] = new VarP();

            AddHiddenValueHandlingFunction(new Average(), 101);
            AddHiddenValueHandlingFunction(new Count(), 102);
            AddHiddenValueHandlingFunction(new CountA(), 103);
            AddHiddenValueHandlingFunction(new Max(), 104);
            AddHiddenValueHandlingFunction(new Min(), 105);
            AddHiddenValueHandlingFunction(new Product(), 106);
            AddHiddenValueHandlingFunction(new Stdev(), 107);
            AddHiddenValueHandlingFunction(new StdevP(), 108);
            AddHiddenValueHandlingFunction(new SumSubtotal(), 109);
            AddHiddenValueHandlingFunction(new Var(), 110);
            AddHiddenValueHandlingFunction(new VarP(), 111);
        }

        private void AddHiddenValueHandlingFunction(HiddenValuesHandlingFunction func, int funcNum)
        {
            func.IgnoreHiddenValues = true;
            _functions[funcNum] = func;
        }

        public override int ArgumentMinLength => 2;
        public override CompileResult Execute(IList<FunctionArgument> arguments, ParsingContext context)
        {
            var funcNum = ArgToInt(arguments, 0);

            var cellId = ExcelCellBase.GetCellId(context.CurrentCell.WorksheetIx, context.CurrentCell.Row, context.CurrentCell.Column);
            if (!context.SubtotalAddresses.Contains(cellId))
            {
                context.SubtotalAddresses.Add(cellId);
            }
            context.IsSubtotal = true;

            var actualArgs = arguments.Skip(1);
            var function = GetFunctionByCalcType(funcNum);
            var compileResult = function.Execute(actualArgs.ToList(), context);
            compileResult.IsResultOfSubtotal = true;
            return compileResult;
        }

        private ExcelFunction GetFunctionByCalcType(int funcNum)
        {
            if (!_functions.ContainsKey(funcNum))
            {
                ThrowExcelErrorValueException(eErrorType.Value);
                //throw new ArgumentException("Invalid funcNum " + funcNum + ", valid ranges are 1-11 and 101-111");
            }
            return _functions[funcNum];
        }
    }
}
