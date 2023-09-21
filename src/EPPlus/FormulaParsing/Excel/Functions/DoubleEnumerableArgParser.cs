/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  03/10/2023         EPPlus Software AB       Initial release EPPlus 7
 *************************************************************************************************/
using OfficeOpenXml.Utils;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions
{
    internal class DoubleEnumerableArgParser
    {
        public DoubleEnumerableArgParser(
            IEnumerable<FunctionArgument> args, 
            ParsingContext context,
            DoubleEnumerableParseOptions options)
        {
            _args = args;
            _context = context;
            _options = options;
            _result = new List<double>();
        }

        private readonly IEnumerable<FunctionArgument> _args;
        private readonly ParsingContext _context;
        private readonly DoubleEnumerableParseOptions _options;
        private readonly List<double> _result;

        private void Parse(IEnumerable<FunctionArgument> args, DoubleEnumerableParseOptions options, out ExcelErrorValue error)
        {
            error = null;
            foreach (var arg in args)
            {
                if (arg.IsExcelRange)
                {
                    foreach (var cell in arg.ValueAsRangeInfo)
                    {
                        if (!options.IgnoreErrors && cell.IsExcelError) 
                        {
                            error = ExcelErrorValue.Parse(cell.Value.ToString());
                            return;
                        }
                        if (!CellStateHelper.ShouldIgnore(options.IgnoreHiddenCells, options.IgnoreNonNumeric, cell, _context, options.IgnoreNestedSubtotalAggregate) && ConvertUtil.IsExcelNumeric(cell.Value))
                        {
                            var val = new ExcelDoubleCellValue(cell.ValueDouble, cell.Row, cell.Column);
                            _result.Add(val);
                        }
                    }
                }
                else if(arg.Value is IEnumerable<FunctionArgument> fArgs)
                {
                    Parse(fArgs, options, out error);
                }
                else
                {
                    if (!options.IgnoreErrors && arg.ValueIsExcelError)
                    {
                        error = arg.ValueAsExcelErrorValue;
                        return;
                    }
                    if (ConvertUtil.IsExcelNumeric(arg.Value) && !CellStateHelper.ShouldIgnore(options.IgnoreHiddenCells, options.IgnoreNestedSubtotalAggregate, arg, _context))
                    {
                        var val = new ExcelDoubleCellValue(ConvertUtil.GetValueDouble(arg.Value));
                        _result.Add(val);
                    }
                }
            }
        }

        public IList<double> GetResult(out ExcelErrorValue error)
        {
            Parse(_args, _options, out error);
            if(error != null)
            {
                return new List<double>();
            }
            return _result;
        }
    }
}
