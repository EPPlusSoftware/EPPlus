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

        public virtual IEnumerable<ExcelDoubleCellValue> ArgsToDoubleEnumerable(bool ignoreHidden, bool ignoreErrors, bool ignoreSubtotalAggregate, IEnumerable<FunctionArgument> arguments, ParsingContext context, bool ignoreNonNumeric = false)
        {
            return _doubleEnumerableArgConverter.ConvertArgs(ignoreHidden, ignoreErrors, ignoreSubtotalAggregate, arguments, context, ignoreNonNumeric);
        }

        public virtual IEnumerable<object> ArgsToObjectEnumerable(bool ignoreHidden,
                                                                  bool ignoreErrors,
                                                                  bool ignoreNestedSubtotalAggregate,
                                                                  IEnumerable<FunctionArgument> arguments,
                                                                  ParsingContext context)
        {
            return _objectEnumerableArgConverter.ConvertArgs(ignoreHidden, ignoreErrors, ignoreNestedSubtotalAggregate, arguments, context);
        }

        public virtual IEnumerable<object> ArgsToObjectEnumerable(bool ignoreHidden,
                                                                  bool ignoreErrors,
                                                                  IEnumerable<FunctionArgument> arguments,
                                                                  ParsingContext context)
        {
            return _objectEnumerableArgConverter.ConvertArgs(ignoreHidden, ignoreErrors, arguments, context);
        }
    }
}
