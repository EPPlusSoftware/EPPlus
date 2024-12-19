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
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;

namespace OfficeOpenXml.FormulaParsing.FormulaExpressions
{
    /// <summary>
    /// Indicates that the result the function should be created as a dynamic array result.
    /// </summary>
    /// <summary>
    /// Indicates that the result the function should be created as a dynamic array result.
    /// </summary>
    public class DynamicArrayCompileResult : AddressCompileResult
    {
        CompileResultType _resultType = CompileResultType.DynamicArray;
        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="result"></param>
        /// <param name="dataType"></param>
        /// <param name="address"></param>
        /// <param name="resultType"></param>
        public DynamicArrayCompileResult(object result, DataType dataType, FormulaRangeAddress address, CompileResultType resultType) : base(result, dataType, address)
        {
            _resultType = resultType;
        }

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="result"></param>
        /// <param name="dataType"></param>
        /// <param name="address"></param>
        public DynamicArrayCompileResult(object result, DataType dataType, FormulaRangeAddress address) : base(result, dataType, address)
        {

        }
        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="result"></param>
        /// <param name="dataType"></param>
        public DynamicArrayCompileResult(object result, DataType dataType) : base(result, dataType)
        {

        }
        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="error"></param>
        public DynamicArrayCompileResult(eErrorType error) : base(error)
        {

        }
        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="errorValue"></param>
        public DynamicArrayCompileResult(ExcelErrorValue errorValue) : base(errorValue)
        {

        }
        /// <summary>
        /// The result is a dynamic array.
        /// </summary>
        public override CompileResultType ResultType
        {
            get
            {
                return _resultType;
            }
        }
    }
}
