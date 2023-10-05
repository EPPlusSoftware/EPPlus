﻿/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  04/03/2020         EPPlus Software AB           EPPlus 5.1
 *************************************************************************************************/
using OfficeOpenXml.FormulaParsing.FormulaExpressions;
using OfficeOpenXml.Utils;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.MathFunctions
{
    internal abstract class SumxBase : ExcelFunction
    {
        private ParsingContext _context;
        public override CompileResult Execute(IList<FunctionArgument> arguments, ParsingContext context)
        {
            _context = context;
            var arg1 = arguments[0];
            var arg2 = arguments[1];
            CreateSets(arg1, arg2, out double[] set1, out double[] set2, out ExcelErrorValue e1);
            if (e1 != null) return CompileResult.GetErrorResult(e1.Type);
            if (set1.Length != set2.Length) return CreateResult(eErrorType.NA);
            var result = Calculate(set1.ToArray(), set2.ToArray());
            return CreateResult(result, DataType.Decimal);
        }

        public abstract double Calculate(double[] set1, double[] set2);

        private void CreateSets(FunctionArgument arg1, FunctionArgument arg2, out double[] set1, out double[] set2, out ExcelErrorValue error)
        {
            set1 = new double[0];
            set2 = new double[0];
            var list1 = CreateSet(arg1, out error);
            if (error != null) return;
            var list2 = CreateSet(arg2, out error);
            if(error != null) return;
            if(list1.Count == list2.Count)
            {
                var r1 = new List<double>();
                var r2 = new List<double>();
                for(var x = 0; x < list1.Count; x++)
                {
                    if(!double.IsNaN(list1[x]) && !double.IsNaN(list2[x]))
                    {
                        r1.Add(list1[x]);
                        r2.Add(list2[x]);
                    }
                }
                set1 = r1.ToArray();
                set2 = r2.ToArray();
            }
            else
            {
                set1 = list1.ToArray();
                set2 = list2.ToArray();
            }
        }

        public List<double> CreateSet(FunctionArgument arg, out ExcelErrorValue error)
        {
            error = null;
            List<double> result = new List<double>();
            if (arg.IsExcelRange)
            {
                var r1 = arg.ValueAsRangeInfo;
                for (var x = 0; x < r1.Count(); x++)
                {
                    var v = r1.ElementAt(x).Value;
                    if (!IsNumeric(v))
                    {
                        result.Add(double.NaN);
                    }
                    else
                    {
                        result.Add(ConvertUtil.GetValueDouble(v));
                    }
                }
            }
            else
            {
                result = ArgsToDoubleEnumerable(arg, _context, out error).ToList();
            }
            return result;
        }
    }
}
