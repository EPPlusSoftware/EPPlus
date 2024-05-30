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
using System.Globalization;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using OfficeOpenXml.Utils;
using OfficeOpenXml.FormulaParsing.Excel.Functions;
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
using OfficeOpenXml.FormulaParsing.Ranges;
using OfficeOpenXml.FormulaParsing.Excel.Operators;
using OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Information;

namespace OfficeOpenXml.FormulaParsing.FormulaExpressions
{
    /// <summary>
    /// Result type
    /// </summary>
    public enum CompileResultType
    {
        /// <summary>
        /// A normal compile result containing a value.
        /// </summary>
        Normal = 0,
        /// <summary>
        /// A compile result referencing a range address. This will allow the result to be used with the colon operator.
        /// </summary>
        RangeAddress = 1,
        /// <summary>
        /// The result is a dynamic array formula.
        /// </summary>
        DynamicArray = 2
    }
    /// <summary>
    /// CompileResultBase
    /// </summary>
    public abstract class CompileResultBase
    {
        /// <summary>
        /// Result type
        /// </summary>
        public abstract CompileResultType ResultType { get; }
    }
    /// <summary>
    /// Compile result
    /// </summary>
    public class CompileResult : CompileResultBase
    {
        private static CompileResult _empty = new CompileResult(null, DataType.Empty);
        private static CompileResult _zeroDecimal = new CompileResult(0d, DataType.Decimal);
        private static CompileResult _zeroInt = new CompileResult(0d, DataType.Integer);


        private static CompileResult _errorRef = new CompileResult(ErrorValues.RefError, DataType.ExcelError);
        private static CompileResult _errorValue = new CompileResult(ErrorValues.ValueError, DataType.ExcelError);
        private static CompileResult _errorNA = new CompileResult(ErrorValues.NAError, DataType.ExcelError);
        private static CompileResult _errorDiv0 = new CompileResult(ErrorValues.Div0Error, DataType.ExcelError);
        private static CompileResult _errorNull = new CompileResult(ErrorValues.NullError, DataType.ExcelError);
        private static CompileResult _errorName = new CompileResult(ErrorValues.NameError, DataType.ExcelError);
        private static CompileResult _errorNum = new CompileResult(ErrorValues.NumError, DataType.ExcelError);
        private static CompileResult _errorCalc = new CompileResult(ErrorValues.CalcError, DataType.ExcelError);

        private static DynamicArrayCompileResult _arrayErrorRef = new DynamicArrayCompileResult(ErrorValues.RefError, DataType.ExcelError);
        private static DynamicArrayCompileResult _arrayErrorValue = new DynamicArrayCompileResult(ErrorValues.ValueError, DataType.ExcelError);
        private static DynamicArrayCompileResult _arrayErrorNA = new DynamicArrayCompileResult(ErrorValues.NAError, DataType.ExcelError);
        private static DynamicArrayCompileResult _arrayErrorDiv0 = new DynamicArrayCompileResult(ErrorValues.Div0Error, DataType.ExcelError);
        private static DynamicArrayCompileResult _arrayErrorNull = new DynamicArrayCompileResult(ErrorValues.NullError, DataType.ExcelError);
        private static DynamicArrayCompileResult _arrayErrorName = new DynamicArrayCompileResult(ErrorValues.NameError, DataType.ExcelError);
        private static DynamicArrayCompileResult _arrayErrorNum = new DynamicArrayCompileResult(ErrorValues.NumError, DataType.ExcelError);
        private static DynamicArrayCompileResult _arrayErrorCalc = new DynamicArrayCompileResult(ErrorValues.CalcError, DataType.ExcelError);


        //private static CompileResult _errorSpill = new CompileResult(ErrorValues.SpillError, DataType.ExcelError); //Spill should use the Spill error containing row and column offset.


        /// <summary>
        /// Returns a CompileResult with a null value and data type set to DataType.Empty
        /// </summary>
        public static CompileResult Empty
        {
            get { return _empty; }
        }

        /// <summary>
        /// Returns a CompileResult instance with a decimal value of 0.
        /// </summary>
        public static CompileResult ZeroDecimal
        {
            get { return _zeroDecimal; }
        }

        /// <summary>
        /// Returns a CompileResult instance with a integer value of 0.
        /// </summary>
        public static CompileResult ZeroInt
        {
            get { return _zeroInt; }
        }
        /// <summary>
        /// Returns a CompileResult instance with a boolean value of false.
        /// </summary>
        public static CompileResult False { get; } = new CompileResult(false, DataType.Boolean);
        /// <summary>
        /// Returns a CompileResult instance with a boolean value of true.
        /// </summary>
        public static CompileResult True { get; } = new CompileResult(true, DataType.Boolean);

        private double? _resultNumeric;
        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="result">The result.</param>
        /// <param name="dataType">The data type of the result.</param>
        public CompileResult(object result, DataType dataType)
        {
            if(result is ExcelDoubleCellValue v)
            {
                Result = v.Value;
            }
            else
            {
                Result = result;
            }
            DataType = dataType;
        }

        internal CompileResult Negate()
        {
            if(DataType == DataType.ExcelRange && Result is IRangeInfo ri)
            {
                return new CompileResult(RangeOperationsOperator.Negate(ri), DataType.ExcelRange);
            }

            else if (IsNumeric)
            {
                return new CompileResult(ResultNumeric * -1, DataType.Decimal);
            }
            else if (DataType != DataType.ExcelError)
            {
                return _errorValue;
            }
            return this;
        }
        internal static CompileResult GetDynamicArrayResultError(eErrorType errorType)
        {
            switch (errorType)
            {
                case eErrorType.Ref:
                    return _arrayErrorRef;
                case eErrorType.Name:
                    return _arrayErrorName;
                case eErrorType.Null:
                    return _arrayErrorNull;
                case eErrorType.Div0:
                    return _arrayErrorDiv0;
                case eErrorType.NA:
                    return _arrayErrorNA;
                case eErrorType.Num:
                    return _arrayErrorNum;
                case eErrorType.Calc:
                    return _arrayErrorCalc;
                default: //#Value!
                    return _arrayErrorValue;
            }
        }

        /// <summary>
        /// Returns a <see cref="CompileResult" /> from the error type/>
        /// </summary>
        /// <param name="errorType">The type of error.</param>
        /// <returns>The <see cref="CompileResult" /> with a the value containing the <see cref="ExcelErrorValue"/> for the type.</returns>
        public static CompileResult GetErrorResult(eErrorType errorType)
        {
            switch (errorType)
            {
                case eErrorType.Ref:
                    return _errorRef;
                case eErrorType.Name:
                    return _errorName;
                case eErrorType.Null:
                    return _errorNull;
                case eErrorType.Div0:
                    return _errorDiv0;
                case eErrorType.NA:
                    return _errorNA;
                case eErrorType.Num:
                    return _errorNum;
                case eErrorType.Calc:
                    return _errorCalc;
                default: //#Value!
                    return _errorValue;
            }
        }

        /// <summary>
        /// Compile result with error type
        /// </summary>
        /// <param name="errorType"></param>
        public CompileResult(eErrorType errorType)
        {
            Result = ExcelErrorValue.Create(errorType);
            DataType = DataType.ExcelError;
        }
        /// <summary>
        /// Compile result with error value
        /// </summary>
        /// <param name="errorValue"></param>
        public CompileResult(ExcelErrorValue errorValue)
        {
            Require.Argument(errorValue).IsNotNull("errorValue");
            Result = errorValue;
            DataType = DataType.ExcelError;
        }
        /// <summary>
        /// RESULT
        /// </summary>
        public object Result
        {
            get;
            private set;
        }
        /// <summary>
        /// Result Value
        /// </summary>
        public object ResultValue
        {
            get
            {
                if(DataType==DataType.ExcelRange)
                {
                    var r = Result as IRangeInfo;
                    if (r == null || r.GetNCells()>1)
                    {
                        return Result;
                    }
                    else
                    {
                        return r.GetValue(r.Address.FromRow, r.Address.FromCol);
                    }
                }
                else
                {
                    return Result;
                }
            }
        }
        /// <summary>
        /// Result numeric
        /// </summary>
        public double ResultNumeric
        {
            get
            {
				// We assume that Result does not change unless it is a range.
				if (_resultNumeric == null)
				{
					if (IsNumeric)
					{
						_resultNumeric = Result == null ? 0 : Convert.ToDouble(Result);
					}
                    else if(IsPercentageString && ConvertUtil.TryParsePercentageString(Result.ToString(), out double v))
                    {
                        _resultNumeric = v;
                    }
					else if (Result is DateTime)
					{
                        _resultNumeric = ((DateTime)Result).ToOADate();
					}
					else if (Result is TimeSpan)
					{
                        _resultNumeric = DateTime.FromOADate(0).Add((TimeSpan)Result).ToOADate();
					}
					else if (Result is IRangeInfo)
					{
						var c = ((IRangeInfo)Result).FirstOrDefault();
						if (c == null)
						{
							return 0;
						}
						else
						{
							return c.ValueDoubleLogical;
						}
					}
					// The IsNumericString and IsDateString properties will set _resultNumeric for efficiency so we just need
					// to check them here.
					else if (!IsDateString && !IsNumericString)
					{
						_resultNumeric = 0;
					}
				}
				return _resultNumeric.Value;
            }
        }
        /// <summary>
        /// Data type
        /// </summary>
        public DataType DataType
        {
            get;
            private set;
        }
        /// <summary>
        /// Is the result numeric
        /// </summary>
        public bool IsNumeric
        {
            get 
            {
                return DataType == DataType.Decimal || DataType == DataType.Integer || DataType == DataType.Empty || DataType == DataType.Boolean || DataType == DataType.Date || DataType == DataType.Time;
            }
        }

        /// <summary>
        /// Is result numeric string
        /// </summary>
        public bool IsNumericString
        {
            get
            {
                if (DataType == DataType.String && ConvertUtil.TryParseNumericString(Result as string, out double result))
                {
                    _resultNumeric = result;
                    return true;
                }
                return false;
            }
        }
        /// <summary>
        /// Is percentage string
        /// </summary>
        public bool IsPercentageString
        {
            get
            {
                if (DataType == DataType.String)
                {
                    var s = Result as string;
                    return ConvertUtil.IsPercentageString(s);
                }
                return false;
            }
            
        }
        /// <summary>
        /// Is date string
        /// </summary>
		public bool IsDateString
		{
			get
			{
                if (DataType == DataType.String && ConvertUtil.TryParseDateString((Result as string), out DateTime result))
                {
                    _resultNumeric = result.ToOADate();
                    return true;
                }
                return false;
			}
		}
        /// <summary>
        /// Is result of subtotal
        /// </summary>
		public bool IsResultOfSubtotal { get; set; }

        /// <summary>
        /// Is hidden cell
        /// </summary>
        public bool IsHiddenCell { get; set; }

        //public int ExcelAddressReferenceId { get; set; }

        /// <summary>
        /// Is result of resolved excelRange
        /// </summary>
        public bool IsResultOfResolvedExcelRange
        {
            get { return Address != null; }
        }
        /// <summary>
        /// Range address
        /// </summary>
        public virtual FormulaRangeAddress Address
        {
            get
            {
                return null;
            }
        }
        /// <summary>
        /// Result type
        /// </summary>
        public override CompileResultType ResultType
        {
            get
            {
                return CompileResultType.Normal;
            }
        }

    }
    /// <summary>
    /// Address compile result
    /// </summary>
    public class AddressCompileResult : CompileResult
    {
        /// <summary>
        /// Address result
        /// </summary>
        /// <param name="result"></param>
        /// <param name="dataType"></param>
        /// <param name="address"></param>
        public AddressCompileResult(object result, DataType dataType, FormulaRangeAddress address) : base(result, dataType)
        {
            Address = address;
        }
        /// <summary>
        /// Address result without address
        /// </summary>
        /// <param name="result"></param>
        /// <param name="dataType"></param>
        public AddressCompileResult(object result, DataType dataType) : base(result, dataType)
        { 

        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="error"></param>
        public AddressCompileResult(eErrorType error) : base(error)
        {

        }
        /// <summary>
        /// Address compile result
        /// </summary>
        /// <param name="errorValue"></param>
        public AddressCompileResult(ExcelErrorValue errorValue) : base(errorValue)
        {

        }
        /// <summary>
        /// Address
        /// </summary>
        public override FormulaRangeAddress Address
        {
            get;
        }
        /// <summary>
        /// ResultType
        /// </summary>
        public override CompileResultType ResultType
        {
            get
            {
                if(Address==null)
                {
                    return base.ResultType;
                }
                return CompileResultType.RangeAddress;
            }
        }
    }
    /// <summary>
    /// Indicates that the result the function should be created as a dynamic array result.
    /// </summary>
    public class DynamicArrayCompileResult : AddressCompileResult
    {
        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="result"></param>
        /// <param name="dataType"></param>
        /// <param name="address"></param>
        public DynamicArrayCompileResult(object result, DataType dataType, FormulaRangeAddress address) : base(result, dataType)
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
                return CompileResultType.DynamicArray;
            }
        }
    }
}
