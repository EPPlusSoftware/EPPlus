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
    /// <summary>
    /// Simplifies function argument input by collecting and enumerating arguments of different types
    /// </summary>
    internal class ArgumentCollectionUtil
    {
        private readonly DoubleEnumerableArgConverter _doubleEnumerableArgConverter;
        private readonly ObjectEnumerableArgConverter _objectEnumerableArgConverter;

        /// <summary>
        /// Empty constructor
        /// </summary>
        public ArgumentCollectionUtil()
            : this(new DoubleEnumerableArgConverter(), new ObjectEnumerableArgConverter())
        {

        }

        /// <summary>
        /// Constructor with converters
        /// </summary>
        /// <param name="doubleEnumerableArgConverter"></param>
        /// <param name="objectEnumerableArgConverter"></param>
        public ArgumentCollectionUtil(
            DoubleEnumerableArgConverter doubleEnumerableArgConverter, 
            ObjectEnumerableArgConverter objectEnumerableArgConverter)
        {
            _doubleEnumerableArgConverter = doubleEnumerableArgConverter;
            _objectEnumerableArgConverter = objectEnumerableArgConverter;
        }

        /// <summary>
        /// Converts args to enumerable ExcelDoubleCellValue
        /// </summary>
        /// <param name="ignoreHidden"></param>
        /// <param name="ignoreErrors"></param>
        /// <param name="ignoreSubtotalAggregate"></param>
        /// <param name="arguments"></param>
        /// <param name="context"></param>
        /// <param name="ignoreNonNumeric"></param>
        /// <returns></returns>
        public virtual IEnumerable<ExcelDoubleCellValue> ArgsToDoubleEnumerable(bool ignoreHidden, bool ignoreErrors, bool ignoreSubtotalAggregate, IEnumerable<FunctionArgument> arguments, ParsingContext context, bool ignoreNonNumeric = false)
        {
            return _doubleEnumerableArgConverter.ConvertArgs(ignoreHidden, ignoreErrors, ignoreSubtotalAggregate, arguments, context, ignoreNonNumeric);
        }

        /// <summary>
        /// Converts args to enumerable objects with an aggregate
        /// </summary>
        /// <param name="ignoreHidden"></param>
        /// <param name="ignoreErrors"></param>
        /// <param name="ignoreNestedSubtotalAggregate"></param>
        /// <param name="arguments"></param>
        /// <param name="context"></param>
        /// <returns></returns>
        public virtual IEnumerable<object> ArgsToObjectEnumerable(bool ignoreHidden,
                                                                  bool ignoreErrors,
                                                                  bool ignoreNestedSubtotalAggregate,
                                                                  IEnumerable<FunctionArgument> arguments,
                                                                  ParsingContext context)
        {
            return _objectEnumerableArgConverter.ConvertArgs(ignoreHidden, ignoreErrors, ignoreNestedSubtotalAggregate, arguments, context);
        }

        /// <summary>
        /// Converts args to enumerable objects
        /// </summary>
        /// <param name="ignoreHidden"></param>
        /// <param name="ignoreErrors"></param>
        /// <param name="arguments"></param>
        /// <param name="context"></param>
        /// <returns></returns>
        public virtual IEnumerable<object> ArgsToObjectEnumerable(bool ignoreHidden,
                                                                  bool ignoreErrors,
                                                                  IEnumerable<FunctionArgument> arguments,
                                                                  ParsingContext context)
        {
            return _objectEnumerableArgConverter.ConvertArgs(ignoreHidden, ignoreErrors, arguments, context);
        }
    }
}
