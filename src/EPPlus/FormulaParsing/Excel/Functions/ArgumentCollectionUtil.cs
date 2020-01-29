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

namespace OfficeOpenXml.FormulaParsing.Excel.Functions
{
    public class ArgumentCollectionUtil
    {
        private readonly DoubleEnumerableArgConverter _doubleEnumerableArgConverter;
        private readonly ObjectEnumerableArgConverter _objectEnumerableArgConverter;

        public ArgumentCollectionUtil()
            : this(new DoubleEnumerableArgConverter(), new ObjectEnumerableArgConverter())
        {

        }

        public ArgumentCollectionUtil(
            DoubleEnumerableArgConverter doubleEnumerableArgConverter, 
            ObjectEnumerableArgConverter objectEnumerableArgConverter)
        {
            _doubleEnumerableArgConverter = doubleEnumerableArgConverter;
            _objectEnumerableArgConverter = objectEnumerableArgConverter;
        }

        public virtual IEnumerable<ExcelDoubleCellValue> ArgsToDoubleEnumerable(bool ignoreHidden, bool ignoreErrors, IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            return _doubleEnumerableArgConverter.ConvertArgs(ignoreHidden, ignoreErrors, arguments, context);
        }

        public virtual IEnumerable<object> ArgsToObjectEnumerable(bool ignoreHidden,
                                                                  IEnumerable<FunctionArgument> arguments,
                                                                  ParsingContext context)
        {
            return _objectEnumerableArgConverter.ConvertArgs(ignoreHidden, arguments, context);
        }

        public virtual double CalculateCollection(IEnumerable<FunctionArgument> collection, double result, Func<FunctionArgument, double, double> action)
        {
            foreach (var item in collection)
            {
                if (item.Value is IEnumerable<FunctionArgument>)
                {
                    result = CalculateCollection((IEnumerable<FunctionArgument>)item.Value, result, action);
                }
                else
                {
                    result = action(item, result);
                }
            }
            return result;
        }
    }
}
