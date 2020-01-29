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
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Math
{
    internal class Rank : ExcelFunction
    {
        bool _isAvg;
        public Rank(bool isAvg=false)
        {
            _isAvg=isAvg;
        }
        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            ValidateArguments(arguments, 2);
            var number = ArgToDecimal(arguments, 0);
            var refer = arguments.ElementAt(1);
            bool asc = false;
            if (arguments.Count() > 2)
            {
                asc = base.ArgToBool(arguments, 2);
            }
            var l = new List<double>();

            foreach (var c in refer.ValueAsRangeInfo)
            {
                var v = Utils.ConvertUtil.GetValueDouble(c.Value, false, true);
                if (!double.IsNaN(v))
                {
                    l.Add(v);
                }
            }
            l.Sort();
            double ix;
            if (asc)
            {
                ix = l.IndexOf(number)+1;
                if(_isAvg)
                {
                    int st = Convert.ToInt32(ix);
                    while (l.Count > st && l[st] == number) st++;
                    if (st > ix) ix = ix + ((st - ix) / 2D);
                }
            }
            else
            {
                ix = l.LastIndexOf(number);
                if (_isAvg)
                {
                    int st = Convert.ToInt32(ix)-1;
                    while (0 <= st && l[st] == number) st--;
                    if (st+1 < ix) ix = ix - ((ix - st - 1) / 2D);
                }
                ix = l.Count - ix;
            }
            if (ix <= 0 || ix>l.Count)
            {
                return new CompileResult(ExcelErrorValue.Create(eErrorType.NA), DataType.ExcelError);
            }
            else
            {
                return CreateResult(ix, DataType.Decimal);
            }
        }
    }
}
