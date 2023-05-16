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

namespace OfficeOpenXml.FormulaParsing.FormulaExpressions
{
    public class CompileResult
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
        //private static CompileResult _errorSpill = new CompileResult(ErrorValues.SpillError, DataType.ExcelError);


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

        private double? _resultNumeric;

        public CompileResult(object result, DataType dataType)
            : this(result, dataType, 0)
        { 
        }

        public CompileResult(object result, DataType dataType, int excelAddressReferenceId)
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
            //ExcelAddressReferenceId = excelAddressReferenceId;
        }

        internal void Negate()
        {
            if(DataType == DataType.ExcelRange && Result is IRangeInfo ri)
            {
                Result = RangeOperationsOperator.Negate(ri);
            }

            else if (ResultNumeric != 0)
            {
                _resultNumeric *= -1;
                Result = _resultNumeric;
            }
        }
        internal static CompileResult GetErrorResult(eErrorType errorType)
        {
            switch(errorType)
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

        public CompileResult(eErrorType errorType)
        {
            Result = ExcelErrorValue.Create(errorType);
            DataType = DataType.ExcelError;
        }

        public CompileResult(ExcelErrorValue errorValue)
        {
            Require.Argument(errorValue).IsNotNull("errorValue");
            Result = errorValue;
            DataType = DataType.ExcelError;
        }

        public object Result
        {
            get;
            private set;
        }

        public object ResultValue
        {
            get
            {
                if(DataType==DataType.ExcelRange)
                {
                    var r = Result as IRangeInfo;
                    if (r == null || r.IsMulti)
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

        public DataType DataType
        {
            get;
            private set;
        }
        
        public bool IsNumeric
        {
            get 
            {
                return DataType == DataType.Decimal || DataType == DataType.Integer || DataType == DataType.Empty || DataType == DataType.Boolean || DataType == DataType.Date || DataType == DataType.Time;
            }
        }

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

		public bool IsResultOfSubtotal { get; set; }

        public bool IsHiddenCell { get; set; }

        //public int ExcelAddressReferenceId { get; set; }

        public bool IsResultOfResolvedExcelRange
        {
            get { return Address != null; }
        }
        public virtual FormulaRangeAddress Address
        {
            get
            {
                return null;
            }
        }
    }
    public class AddressCompileResult : CompileResult
    {
        public AddressCompileResult(Object result, DataType dataType, FormulaRangeAddress address) : base(result, dataType)
        {
            Address = address;
        }
        public AddressCompileResult(Object result, DataType dataType) : base(result, dataType)
        { 

        }
        public AddressCompileResult(eErrorType error) : base(error)
        {

        }
        public AddressCompileResult(ExcelErrorValue errorValue) : base(errorValue)
        {

        }
        public override FormulaRangeAddress Address
        {
            get;
        }
    }
}
