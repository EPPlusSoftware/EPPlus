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
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.Exceptions;
using OfficeOpenXml.FormulaParsing.FormulaExpressions;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions
{
    /// <summary>
    /// Base class for functions that needs to handle cells that is not visible.
    /// </summary>
    public abstract class HiddenValuesHandlingFunction : ExcelFunction
    {
        /// <summary>
        /// Hidden values handling function
        /// </summary>
        public HiddenValuesHandlingFunction()
        {
            IgnoreErrors = true;
            IgnoreNestedSubtotalsAndAggregates = true;
        }
        /// <summary>
        /// Set to true or false to indicate whether the function should ignore hidden values.
        /// </summary>
        public bool IgnoreHiddenValues
        {
            get;
            set;
        }

        /// <summary>
        /// Set to true to indicate whether the function should ignore error values
        /// </summary>
        public bool IgnoreErrors
        {
            get; set;
        }

        /// <summary>
        /// Set to true to indicate whether the function should ignore nested SUBTOTAL and AGGREGATE functions
        /// </summary>
        public bool IgnoreNestedSubtotalsAndAggregates { get; set; }
        /// <summary>
        /// Args to double enumerable
        /// </summary>
        /// <param name="arguments"></param>
        /// <param name="context"></param>
        /// <param name="error"></param>
        /// <returns></returns>
        protected override IList<double> ArgsToDoubleEnumerable(IEnumerable<FunctionArgument> arguments, ParsingContext context, out ExcelErrorValue error)
        {
            return ArgsToDoubleEnumerable(arguments, context, IgnoreErrors, false, out error);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="arguments"></param>
        /// <param name="context"></param>
        /// <param name="ignoreErrors"></param>
        /// <param name="ignoreNonNumeric"></param>
        /// <param name="error"></param>
        /// <returns></returns>
        /// <exception cref="ExcelErrorValueException"></exception>
        protected IList<double> ArgsToDoubleEnumerable(IEnumerable<FunctionArgument> arguments, ParsingContext context, bool ignoreErrors, bool ignoreNonNumeric, out ExcelErrorValue error)
        {
            if (!arguments.Any())
            {
                error = null;
                return new List<double>();
            }
            if (IgnoreHiddenValues)
            {
                var nonHidden = arguments.Where(x => !x.IsHiddenCell);
                var res = base.ArgsToDoubleEnumerable(nonHidden, context, x => x.IgnoreHiddenCells = IgnoreHiddenValues, out ExcelErrorValue e1);
                if (e1 != null) throw new ExcelErrorValueException(e1.Type);
            }
            //return base.ArgsToDoubleEnumerable(IgnoreHiddenValues, ignoreErrors, IgnoreNestedSubtotalsAndAggregates, arguments, context, ignoreNonNumeric);
            return base.ArgsToDoubleEnumerable(arguments, context, x =>
            {
                x.IgnoreHiddenCells = IgnoreHiddenValues;
                x.IgnoreErrors = ignoreErrors;
                x.IgnoreNestedSubtotalAggregate = IgnoreNestedSubtotalsAndAggregates;
                x.IgnoreNonNumeric = ignoreNonNumeric;
            }, out error);
        }
        /// <summary>
        /// Should Ignore
        /// </summary>
        /// <param name="c"></param>
        /// <param name="context"></param>
        /// <returns></returns>
        protected bool ShouldIgnore(ICellInfo c, ParsingContext context)
        {
            if(CellStateHelper.ShouldIgnore(IgnoreHiddenValues, IgnoreNestedSubtotalsAndAggregates, c, context))
            {
                return true;
            }
            if(IgnoreErrors && c.IsExcelError)
            {
                return true;
            }
            return false;
        }
        /// <summary>
        /// Should ignore with argument
        /// </summary>
        /// <param name="arg"></param>
        /// <param name="context"></param>
        /// <returns></returns>
        protected bool ShouldIgnore(FunctionArgument arg, ParsingContext context)
        {
            if (CellStateHelper.ShouldIgnore(IgnoreHiddenValues, IgnoreNestedSubtotalsAndAggregates, arg, context))
            {
                return true;
            }
            if(IgnoreErrors && arg.ValueIsExcelError)
            {
                return true;
            }
            return false;
        }

    }
}
