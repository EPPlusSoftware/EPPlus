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
using System.Text.RegularExpressions;
using OfficeOpenXml.FormulaParsing.FormulaExpressions;
using System.Globalization;
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.Utilities;
using OfficeOpenXml.FormulaParsing.Exceptions;
using System.Collections;
using static OfficeOpenXml.FormulaParsing.EpplusExcelDataProvider;
using static OfficeOpenXml.FormulaParsing.ExcelDataProvider;
using OfficeOpenXml.Compatibility;
using OfficeOpenXml.FormulaParsing.Excel.Functions.MathFunctions;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
using System.Runtime.CompilerServices;
using Utils = OfficeOpenXml.Utils;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions
{
    /// <summary>
    /// Base class for Excel function implementations.
    /// </summary>
    public abstract class ExcelFunction
    {        
        public ExcelFunction()
            : this(new ArgumentCollectionUtil(), new ArgumentParsers(), new CompileResultValidators())
        {

        }

        public ExcelFunction(
            ArgumentCollectionUtil argumentCollectionUtil,
            ArgumentParsers argumentParsers,
            CompileResultValidators compileResultValidators)
        {
            _argumentCollectionUtil = argumentCollectionUtil;
            _argumentParsers = argumentParsers;
            _compileResultValidators = compileResultValidators;
            _arrayConfig = new ArrayBehaviourConfig();
            ConfigureArrayBehaviour(_arrayConfig);
        }

        private readonly ArgumentCollectionUtil _argumentCollectionUtil;
        protected readonly ArgumentParsers _argumentParsers;
        private readonly CompileResultValidators _compileResultValidators;
        protected readonly int NumberOfSignificantFigures = 15;
        private readonly ArrayBehaviourConfig _arrayConfig;

        /// <summary>
        /// Configuration for paramenters that can be an array. See <see cref="ConfigureArrayBehaviour(ArrayBehaviourConfig)"/>
        /// </summary>
        internal ArrayBehaviourConfig ArrayBehaviourConfig => _arrayConfig;
        

        /// <summary>
        /// 
        /// </summary>
        /// <param name="arguments">Arguments to the function, each argument can contain primitive types, lists or <see cref="IRangeInfo">Excel ranges</see></param>
        /// <param name="context">The <see cref="ParsingContext"/> contains various data that can be useful in functions.</param>
        /// <returns>A <see cref="CompileResult"/> containing the calculated value</returns>
        public abstract CompileResult Execute(IList<FunctionArgument> arguments, ParsingContext context);

        internal CompileResult ExecuteInternal(IList<FunctionArgument> arguments, ParsingContext context)
        {
            context.HiddenCellBehaviour = HiddenCellHandlingCategory.Default;
            if(arguments==null || arguments.Count < ArgumentMinLength)
            {
                return CompileResult.GetErrorResult(eErrorType.Value);
            }
            
            for (int i = 0; i < arguments.Count; i++)
            {
                if (ParametersInfo.HasNormalArguments)
                {
                    if (arguments[i].DataType == DataType.ExcelError)
                    {
                        return CompileResult.GetErrorResult(arguments[i].ValueAsExcelErrorValue.Type);
                    }
                }
                else
                {
                    var pi = ParametersInfo.GetParameterInfo(i);
                    if (arguments[i].DataType == DataType.ExcelError &&
                        (pi & FunctionParameterInformation.IgnoreErrorInPreExecute) != FunctionParameterInformation.IgnoreErrorInPreExecute)
                    {
                        return CompileResult.GetErrorResult(arguments[i].ValueAsExcelErrorValue.Type);
                    }
                }
            }

            return Execute(arguments, context);
        }

        /// <summary>
        /// Returns the minimum arguments for the function. Number of arguments are validated before calling the execute. If lesser arguments are supplied a #VALUE! error will be returned.
        /// </summary>
        public abstract int ArgumentMinLength { get; }

        /// <summary>
        /// If overridden, this method is called before Execute is called.
        /// </summary>
        /// <param name="context"></param>
        public virtual void BeforeInvoke(ParsingContext context) { }
        public virtual void GetNewParameterAddress(IList<CompileResult> args, int index, ref Queue<FormulaRangeAddress> addresses)
        {
            
        }

        public virtual bool IsErrorHandlingFunction
        {
            get
            {
                return false;
            }
        }

        /// <summary>
        /// Describes how the function works with input ranges and returning arrays.
        /// </summary>
        public virtual ExcelFunctionArrayBehaviour ArrayBehaviour
        {
            get
            {
                return ExcelFunctionArrayBehaviour.None;
            }
        }

        /// <summary>
        /// Configures parameters of a function that can be arrays (multi-cell ranges)
        /// even if the function itself treats them as single values.
        /// </summary>
        /// <param name="config"></param>
        public virtual void ConfigureArrayBehaviour(ArrayBehaviourConfig config)
        {
            if(ArrayBehaviour == ExcelFunctionArrayBehaviour.FirstArgCouldBeARange)
            {
                config.SetArrayParameterIndexes(0);
            }
        }

        /// <summary>
        /// Used for some Lookupfunctions to indicate that function arguments should
        /// not be compiled before the function is called.
        /// </summary>
        //public bool SkipArgumentEvaluation { get; set; }
        protected object GetFirstValue(IEnumerable<FunctionArgument> val)
        {
            var arg = ((IEnumerable<FunctionArgument>)val).FirstOrDefault();
            if (arg.Value is IRangeInfo)
            {
                //var r=((ExcelDataProvider.IRangeInfo)arg);
                var r = arg.ValueAsRangeInfo;
                return r.GetValue(r.Address.FromRow, r.Address.FromCol);
            }
            else
            {
                return arg == null ? null : arg.Value;
            }
        }
        /// <summary>
        /// This functions validates that the supplied <paramref name="arguments"/> contains at least
        /// (the value of) <paramref name="minLength"/> elements. If one of the arguments is an
        /// <see cref="IRangeInfo">Excel range</see> the number of cells in
        /// that range will be counted as well.
        /// </summary>
        /// <param name="arguments"></param>
        /// <param name="minLength"></param>
        /// <param name="errorTypeToThrow">The <see cref="eErrorType"/> of the <see cref="ExcelErrorValueException"/> that will be thrown if <paramref name="minLength"/> is not met.</param>
        protected void ValidateArguments(IEnumerable<FunctionArgument> arguments, int minLength,
                                         eErrorType errorTypeToThrow)
        {
            Require.That(arguments).Named("arguments").IsNotNull();
            ThrowExcelErrorValueExceptionIf(() =>
                {
                    var nArgs = 0;
                    if (arguments.Any())
                    {
                        foreach (var arg in arguments)
                        {
                            nArgs++;
                            if (nArgs >= minLength) return false;
                            if (arg.IsExcelRange)
                            {
                                nArgs += arg.ValueAsRangeInfo.GetNCells();
                                if (nArgs >= minLength) return false;
                            }
                        }
                    }
                    return true;
                }, errorTypeToThrow);
        }

        /// <summary>
        /// This functions validates that the supplied <paramref name="arguments"/> contains at least
        /// (the value of) <paramref name="minLength"/> elements. If one of the arguments is an
        /// <see cref="IRangeInfo">Excel range</see> the number of cells in
        /// that range will be counted as well.
        /// </summary>
        /// <param name="arguments"></param>
        /// <param name="minLength"></param>
        /// <exception cref="ArgumentException"></exception>
        protected void ValidateArguments(IEnumerable<FunctionArgument> arguments, int minLength)
        {
            Require.That(arguments).Named("arguments").IsNotNull();
            ThrowArgumentExceptionIf(() =>
                {
                    var nArgs = 0;
                    if (arguments.Any())
                    {
                        foreach (var arg in arguments)
                        {
                            nArgs++;
                            if (nArgs >= minLength) return false;
                            if (arg.IsExcelRange)
                            {
                                nArgs += arg.ValueAsRangeInfo.GetNCells();
                                if (nArgs >= minLength) return false;
                            }
                        }
                    }
                    return true;
                }, "Expecting at least {0} arguments", minLength.ToString());
        }
        protected string ArgToAddress(IList<FunctionArgument> arguments, int index)
        {
            var arg = arguments[index];

            if (arg.Address != null)
            {
                return arg.Address.WorksheetAddress;
            }

            return ArgToString(arguments, index);
        }
        //protected string ArgToAddress(IEnumerable<FunctionArgument> arguments, int index, ParsingContext context)
        //{
        //    var arg = arguments.ElementAt(index);

        //    if(arg.Address !=null)
        //    {
        //        return arg.Address.WorksheetAddress;
        //    }
        //    return ArgToAddress(arguments, index);
        //}

        /// <summary>
        /// Returns the value of the argument att the position of the 0-based index
        /// <paramref name="index"/> as an integer.
        /// </summary>
        /// <param name="arguments"></param>
        /// <param name="index"></param>
        /// <param name="emptyValue">Value returned if datatype is empty</param>
        /// <returns>Value of the argument as an integer.</returns>
        /// <exception cref="ExcelErrorValueException"></exception>
        protected int ArgToInt(IList<FunctionArgument> arguments, int index, int emptyValue=0)
        {
            var arg = arguments[index];
            switch (arg.DataType)
            {
                case DataType.ExcelError:
                    throw new ExcelErrorValueException(arg.ValueAsExcelErrorValue);
                case DataType.Empty:
                    return emptyValue;
                default:
                    var val = arg.ValueFirst;
                    return (int)_argumentParsers.GetParser(DataType.Integer).Parse(val);
            }
        }

        /// <summary>
        /// Returns the value of the argument att the position of the 0-based index
        /// </summary>
        /// <param name="arguments"></param>
        /// <param name="index"></param>
        /// <param name="ignoreErrors">If true an Excel error in the cell will be ignored</param>
        /// <returns>Value of the argument as an integer.</returns>
        /// /// <exception cref="ExcelErrorValueException"></exception>
        protected int ArgToInt(IList<FunctionArgument> arguments, int index, bool ignoreErrors)
        {
            var arg = arguments[index];
            if (arg.ValueIsExcelError && !ignoreErrors)
            {
                throw new ExcelErrorValueException(arg.ValueAsExcelErrorValue.Type);
            }
            else if (arg.DataType == DataType.Empty)
            {
                return 0;
            }
            return (int)_argumentParsers.GetParser(DataType.Integer).Parse(arg.ValueFirst);
        }

        /// <summary>
        /// Returns the value of the argument att the position of the 0-based
        /// <paramref name="index"/> as an integer.
        /// </summary>
        /// <param name="arguments"></param>
        /// <param name="index"></param>
        /// <param name="roundingMethod"></param>
        /// <returns>Value of the argument as an integer.</returns>
        /// <exception cref="ExcelErrorValueException"></exception>
        protected int ArgToInt(IList<FunctionArgument> arguments, int index, RoundingMethod roundingMethod)
        {
            var arg = arguments[index];
            switch (arg.DataType)
            {
                case DataType.ExcelError:                                        
                   throw new ExcelErrorValueException(arg.ValueAsExcelErrorValue);
                case DataType.Empty:
                    return 0;
                default:
                    var val = arg.ValueFirst;
                    return (int)_argumentParsers.GetParser(DataType.Integer).Parse(val, roundingMethod);
            }
        }

        /// <summary>
        /// Returns the value of the argument att the position of the 0-based
        /// <paramref name="index"/> as a string.
        /// </summary>
        /// <param name="arguments"></param>
        /// <param name="index"></param>
        /// <returns>Value of the argument as a string.</returns>
        protected string ArgToString(IList<FunctionArgument> arguments, int index)
        {
            var obj = arguments[index].ValueFirst;
            return obj != null ? obj.ToString() : string.Empty;
        }

        /// <summary>
        /// Returns the value of the argument att the position of the 0-based
        /// </summary>
        /// <param name="obj"></param>
        /// <param name="error">Will be set if the conversion generated an error</param>
        /// <returns>Value of the argument as a double.</returns>
        /// <exception cref="ExcelErrorValueException"></exception>
        protected double ArgToDecimal(object obj, out ExcelErrorValue error)
        {
            return DoubleArgParser.Parse(obj, out error);
        }

        /// <summary>
        /// Returns the value of the argument att the position of the 0-based
        /// </summary>
        /// <param name="obj"></param>
        /// <param name="precisionAndRoundingStrategy">strategy for handling precision and rounding of double values</param>
        /// <returns>Value of the argument as a double.</returns>
        /// <exception cref="ExcelErrorValueException"></exception>
        protected double ArgToDecimal(object obj, PrecisionAndRoundingStrategy precisionAndRoundingStrategy, out ExcelErrorValue error)
        {
            var result = ArgToDecimal(obj, out error);
            if(error != null)
            {
                return double.NaN;
            }
            if (precisionAndRoundingStrategy == PrecisionAndRoundingStrategy.Excel && result != double.NaN)
            {
                result = RoundingHelper.RoundToSignificantFig(result, NumberOfSignificantFigures);
            }
            return result;
        }
        /// <summary>
        /// Returns the value of the argument att the position of the 0-based
        /// <paramref name="index"/> as a <see cref="System.Double"/>.
        /// </summary>
        /// <param name="arguments"></param>
        /// <param name="index"></param>
        /// <param name="precisionAndRoundingStrategy">strategy for handling precision and rounding of double values</param>
        /// <param name="error">Will be set if an error occurs during conversion</param>
        /// <returns>Value of the argument as an integer.</returns>
        /// <exception cref="ExcelErrorValueException"></exception>
        protected double ArgToDecimal(
            IList<FunctionArgument> arguments, 
            int index,
            out ExcelErrorValue error,
            PrecisionAndRoundingStrategy precisionAndRoundingStrategy = PrecisionAndRoundingStrategy.DotNet
            )
        {
            var arg = arguments[index];
            if(arg.DataType == DataType.ExcelError)
            {
                error = arg.ValueAsExcelErrorValue;
                return double.NaN;
            }
            else if(arg.DataType == DataType.Empty)
            {
                error = null;
                return 0D;
            }
            else
            {
                return ArgToDecimal(arg.Value, precisionAndRoundingStrategy, out error);
            }
        }
        /// <summary>
        /// Returns the value of the argument att the position of the 0-based
        /// <paramref name="index"/> as a <see cref="System.Double"/>.
        /// If the the value is null, zero will be returned.
        /// </summary>
        /// <param name="arguments"></param>
        /// <param name="index"></param>
        /// <param name="valueIfNull"></param>
        /// <returns>Value of the argument as an integer.</returns>
        /// <exception cref="ExcelErrorValueException"></exception>
        protected double ArgToDecimal(IList<FunctionArgument> arguments, int index, double valueIfNull, out ExcelErrorValue error)
        {
            error = null;
            var arg = arguments[index];
            if (arg.Value == null)
            {
                return valueIfNull;
            }
            if (arg.ValueIsExcelError)
            {
                throw new ExcelErrorValueException(arg.ValueAsExcelErrorValue);
            }
            return ArgToDecimal(arg.Value, PrecisionAndRoundingStrategy.DotNet, out error);
        }
        /// <summary>
        /// Returns the value as if the 
        /// </summary>
        /// <param name="arg"></param>
        /// <returns></returns>
        protected double? GetDecimalSingleArgument(FunctionArgument arg)
        {
            if (arg.DataType == DataType.Boolean)
            {
                return arg.Address == null ? Utils.ConvertUtil.GetValueDouble(arg.Value) : default;
            }
            else if (arg.DataType == DataType.String || arg.DataType == DataType.Unknown)
            {
                if (arg.Address != null) return default; //If the value reference a cell address, we ignore strings.
                if (Utils.ConvertUtil.TryParseNumericString(arg.Value.ToString(), out double number))
                {
                    return number;
                }
                else if (Utils.ConvertUtil.TryParseDateString(arg.Value.ToString(), out DateTime date))
                {
                    return date.ToOADate();
                }
                else
                {
                    return default;
                }
            }
            else
            {
                return Utils.ConvertUtil.GetValueDouble(arg.Value);
            }

            return default;
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="arguments"></param>
        /// <param name="index"></param>
        /// <returns></returns>
        protected IRangeInfo ArgToRangeInfo(IList<FunctionArgument> arguments, int index)
        {
            if (arguments[index].DataType == DataType.ExcelRange)
            {
                return arguments[index].Value as IRangeInfo;
            }
            return null;
        }

        protected double Divide(double left, double right)
        {
            if (Math.Abs(right - 0d) < double.Epsilon)
            {
                return double.PositiveInfinity;
            }
            return left / right;
        }

        protected bool IsNumericString(object value)
        {
            if (value == null || string.IsNullOrEmpty(value.ToString())) return false;
            return Regex.IsMatch(value.ToString(), @"^[\d]+(\,[\d])?");
        }

        protected bool IsInteger(object n)
        {
            if (!IsNumeric(n)) return false;
            return Convert.ToDouble(n) % 1 == 0;
        }

        /// <summary>
        /// If the argument is a boolean value its value will be returned.
        /// If the argument is an integer value, true will be returned if its
        /// value is not 0, otherwise false.
        /// </summary>
        /// <param name="arguments"></param>
        /// <param name="index"></param>
        /// <returns></returns>
        protected bool ArgToBool(IList<FunctionArgument> arguments, int index)
        {
            var obj = arguments[index].Value ?? string.Empty;
            return (bool)_argumentParsers.GetParser(DataType.Boolean).Parse(obj);
        }

        /// <summary>
        /// Throws an <see cref="ArgumentException"/> if <paramref name="condition"/> evaluates to true.
        /// </summary>
        /// <param name="condition"></param>
        /// <param name="message"></param>
        /// <exception cref="ArgumentException"></exception>
        protected void ThrowArgumentExceptionIf(Func<bool> condition, string message)
        {
            if (condition())
            {
                throw new ArgumentException(message);
            }
        }

        /// <summary>
        /// Throws an <see cref="ArgumentException"/> if <paramref name="condition"/> evaluates to true.
        /// </summary>
        /// <param name="condition"></param>
        /// <param name="message"></param>
        /// <param name="formats">Formats to the message string.</param>
        protected void ThrowArgumentExceptionIf(Func<bool> condition, string message, params object[] formats)
        {
            message = string.Format(message, formats);
            ThrowArgumentExceptionIf(condition, message);
        }

        /// <summary>
        /// Throws an <see cref="ExcelErrorValueException"/> with the given <paramref name="errorType"/> set.
        /// </summary>
        /// <param name="errorType"></param>
        protected void ThrowExcelErrorValueException(eErrorType errorType)
        {
            throw new ExcelErrorValueException("An excel function error occurred", ExcelErrorValue.Create(errorType));
        }
        /// <summary>
        /// Throws an <see cref="ExcelErrorValueException"/> with the type of given <paramref name="value"/> set.
        /// </summary>
        /// <param name="value"></param>
        protected void ThrowExcelErrorValueException(ExcelErrorValue value)
        {
            if (value != null) throw new ExcelErrorValueException(value.Type);
        }

        /// <summary>
        /// Throws an <see cref="ArgumentException"/> if <paramref name="condition"/> evaluates to true.
        /// </summary>
        /// <param name="condition"></param>
        /// <param name="errorType"></param>
        /// <exception cref="ExcelErrorValueException"></exception>
        protected void ThrowExcelErrorValueExceptionIf(Func<bool> condition, eErrorType errorType)
        {
            if (condition())
            {
                throw new ExcelErrorValueException("An excel function error occurred", ExcelErrorValue.Create(errorType));
            }
        }

#if (!NET35)
        [MethodImpl(MethodImplOptions.AggressiveInlining)]
#endif
        protected bool IsNumeric(object val)
        {
            if (val == null) return false;
            return (TypeCompat.IsPrimitive(val) || val is double || val is decimal || val is DateTime || val is TimeSpan);
        }

        protected bool IsBool(object val)
        {
            return val is bool;
        }

        protected bool IsString(object val, bool allowNullOrEmpty = true)
        {
            if (!allowNullOrEmpty)
                return (val is string) && !string.IsNullOrEmpty(val as string);
            return val is string;
        }

        //protected virtual bool IsNumber(object obj)
        //{
        //    if (obj == null) return false;
        //    return (obj is int || obj is double || obj is short || obj is decimal || obj is long);
        //}

        /// <summary>
        /// Helper method for comparison of two doubles.
        /// </summary>
        /// <param name="d1"></param>
        /// <param name="d2"></param>
        /// <returns></returns>
        protected bool AreEqual(double d1, double d2)
        {
            return System.Math.Abs(d1 - d2) < double.Epsilon;
        }

        /// <summary>
        /// Will return the arguments as an enumerable of doubles.
        /// </summary>
        /// <param name="arguments"></param>
        /// <param name="context"></param>
        /// <returns></returns>
        protected virtual IEnumerable<ExcelDoubleCellValue> ArgsToDoubleEnumerable(IEnumerable<FunctionArgument> arguments,
                                                                     ParsingContext context)
        {
            return ArgsToDoubleEnumerable(false, arguments, context);
        }

        /// <summary>
        /// Will return the arguments as an enumerable of doubles.
        /// </summary>
        /// <param name="ignoreHiddenCells">If a cell is hidden and this value is true the value of that cell will be ignored</param>
        /// <param name="ignoreErrors">If a cell contains an error, that error will be ignored if this method is set to true</param>
        /// <param name="arguments"></param>
        /// <param name="context"></param>
        /// <returns></returns>
        protected virtual IEnumerable<ExcelDoubleCellValue> ArgsToDoubleEnumerable(bool ignoreHiddenCells, bool ignoreErrors, IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            return _argumentCollectionUtil.ArgsToDoubleEnumerable(ignoreHiddenCells, ignoreErrors, false, arguments, context, false);
        }

        /// <summary>
        /// Will return the arguments as an enumerable of doubles.
        /// </summary>
        /// <param name="ignoreHiddenCells">If a cell is hidden and this value is true the value of that cell will be ignored</param>
        /// <param name="ignoreErrors">If a cell contains an error, that error will be ignored if this method is set to true</param>
        /// <param name="ignoreNestedSubtotalAggregate">If cells which value comes from the calculation of a SUBTOTAL or an AGGREGATE function should be ignored, set this to true</param>
        /// <param name="arguments"></param>
        /// <param name="context"></param>
        /// <returns></returns>
        protected virtual IEnumerable<ExcelDoubleCellValue> ArgsToDoubleEnumerable(bool ignoreHiddenCells, bool ignoreErrors, bool ignoreNestedSubtotalAggregate, IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            return _argumentCollectionUtil.ArgsToDoubleEnumerable(ignoreHiddenCells, ignoreErrors, ignoreNestedSubtotalAggregate, arguments, context, false);
        }


        /// <summary>
        /// Will return the arguments as an enumerable of doubles.
        /// </summary>
        /// <param name="ignoreHiddenCells">If a cell is hidden and this value is true the value of that cell will be ignored</param>
        /// <param name="ignoreErrors">If a cell contains an error, that error will be ignored if this method is set to true</param>
        /// <param name="ignoreNestedSubtotalAggregate">If cells which value comes from the calculation of a SUBTOTAL or an AGGREGATE function should be ignored, set this to true</param>
        /// <param name="arguments"></param>
        /// <param name="context"></param>
        /// <param name="ignoreNonNumeric"></param>
        /// <returns></returns>
        protected virtual IEnumerable<ExcelDoubleCellValue> ArgsToDoubleEnumerable(bool ignoreHiddenCells, bool ignoreErrors, bool ignoreNestedSubtotalAggregate, IEnumerable<FunctionArgument> arguments, ParsingContext context, bool ignoreNonNumeric)
        {
            return _argumentCollectionUtil.ArgsToDoubleEnumerable(ignoreHiddenCells, ignoreErrors, ignoreNestedSubtotalAggregate, arguments, context, ignoreNonNumeric);
        }

        /// <summary>
        /// Will return the arguments as an enumerable of doubles.
        /// </summary>
        /// <param name="ignoreHiddenCells">If a cell is hidden and this value is true the value of that cell will be ignored</param>
        /// <param name="ignoreNestedSubtotalAggregate">If cells which value comes from the calculation of a SUBTOTAL or an AGGREGATE function should be ignored, set this to true</param>
        /// <param name="arguments"></param>
        /// <param name="context"></param>
        /// <param name="ignoreNonNumeric"></param>
        /// <returns></returns>
        protected virtual IEnumerable<ExcelDoubleCellValue> ArgsToDoubleEnumerable(bool ignoreHiddenCells, bool ignoreNestedSubtotalAggregate, IEnumerable<FunctionArgument> arguments, ParsingContext context, bool ignoreNonNumeric)
        {
            return ArgsToDoubleEnumerable(ignoreHiddenCells, true, ignoreNestedSubtotalAggregate, arguments, context, ignoreNonNumeric);
        }

        /// <summary>
        /// Will return the arguments as an enumerable of doubles.
        /// </summary>
        /// <param name="ignoreHiddenCells">If a cell is hidden and this value is true the value of that cell will be ignored</param>
        /// <param name="arguments"></param>
        /// <param name="context"></param>
        /// <param name="ignoreNonNumeric"></param>
        /// <returns></returns>
        protected virtual IEnumerable<ExcelDoubleCellValue> ArgsToDoubleEnumerable(bool ignoreHiddenCells, IEnumerable<FunctionArgument> arguments, ParsingContext context, bool ignoreNonNumeric)
        {
            return ArgsToDoubleEnumerable(ignoreHiddenCells, true, false, arguments, context, ignoreNonNumeric);
        }


        /// <summary>
        /// Will return the arguments as an enumerable of doubles.
        /// </summary>
        /// <param name="ignoreHiddenCells">If a cell is hidden and this value is true the value of that cell will be ignored</param>
        /// <param name="arguments"></param>
        /// <param name="context"></param>        
        /// <returns></returns>
        protected virtual IEnumerable<ExcelDoubleCellValue> ArgsToDoubleEnumerable(bool ignoreHiddenCells, IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            return ArgsToDoubleEnumerable(ignoreHiddenCells, true, arguments, context, false);
        }

        protected virtual IEnumerable<double> ArgsToDoubleEnumerableZeroPadded(bool ignoreHiddenCells, IRangeInfo rangeInfo, ParsingContext context)
        {
            var startRow = rangeInfo.Address.FromRow;
            var endRow = rangeInfo.Address.ToRow > rangeInfo.Worksheet.Dimension._toRow ? rangeInfo.Worksheet.Dimension._toRow : rangeInfo.Address.ToRow;
            var startCol = rangeInfo.Address.FromCol;
            var endCol = rangeInfo.Address.ToCol > rangeInfo.Worksheet.Dimension._toCol ? rangeInfo.Worksheet.Dimension._toCol : rangeInfo.Address.ToCol;
            var horizontal = (startRow == endRow && rangeInfo.Address.FromCol < rangeInfo.Address.ToCol);
            var funcArg = new FunctionArgument(rangeInfo, DataType.ExcelRange);
            var result = ArgsToDoubleEnumerable(ignoreHiddenCells, new List<FunctionArgument> { funcArg }, context);
            var dict = new Dictionary<int, double>();
            result.ToList().ForEach(x => dict.Add(horizontal ? x.CellCol.Value : x.CellRow.Value, x.Value));
            var resultList = new List<double>();
            var from = horizontal ? startCol : startRow;
            var to = horizontal ? endCol : endRow;
            for (var row = from; row <= to; row++)
            {
                if (dict.ContainsKey(row))
                {
                    resultList.Add(dict[row]);
                }
                else
                {
                    resultList.Add(0d);
                }
            }
            return resultList;
        }

        /// <summary>
        /// Will return the arguments as an enumerable of objects.
        /// </summary>
        /// <param name="ignoreHiddenCells">If a cell is hidden and this value is true the value of that cell will be ignored</param>
        /// <param name="arguments"></param>
        /// <param name="context"></param>
        /// <returns></returns>
        protected virtual IEnumerable<object> ArgsToObjectEnumerable(bool ignoreHiddenCells, bool ignoreErrors, bool ignoreNestedSubtotalAggregate, IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            return _argumentCollectionUtil.ArgsToObjectEnumerable(ignoreHiddenCells, ignoreErrors, ignoreNestedSubtotalAggregate, arguments, context);
        }

        /// <summary>
        /// Use this method to create a result to return from Excel functions. 
        /// </summary>
        /// <param name="result"></param>
        /// <param name="dataType"></param>
        /// <returns></returns>
        protected CompileResult CreateResult(object result, DataType dataType)
        {
            var validator = _compileResultValidators.GetValidator(dataType);
            validator.Validate(result);
            return new CompileResult(result, dataType);
        }
        protected CompileResult CreateResult(object result, DataType dataType, FormulaRangeAddress address)
        {
            var validator = _compileResultValidators.GetValidator(dataType);
            validator.Validate(result);
            return new AddressCompileResult(result, dataType, address);
        }
        /// <summary>
        /// Use this method to create a result to return from Excel functions. 
        /// </summary>
        /// <param name="result"></param>
        /// <param name="dataType"></param>
        /// <returns></returns>
        protected CompileResult CreateDynamicArrayResult(object result, DataType dataType)
        {
            var validator = _compileResultValidators.GetValidator(dataType);
            validator.Validate(result);
            return new DynamicArrayCompileResult(result, dataType);
        }
        protected CompileResult CreateDynamicArrayResult(object result, DataType dataType, FormulaRangeAddress address)
        {
            var validator = _compileResultValidators.GetValidator(dataType);
            validator.Validate(result);
            return new DynamicArrayCompileResult(result, dataType, address);
        }
        /// <summary>
        /// Use this method to create a result to return from Excel functions. 
        /// </summary>
        /// <param name="result"></param>
        /// <param name="dataType"></param>
        /// <param name="address">The address for the range</param>
        /// <returns></returns>
        protected CompileResult CreateAddressResult(IRangeInfo result, DataType dataType)
        {
            var validator = _compileResultValidators.GetValidator(dataType);
            validator.Validate(result);
            return new AddressCompileResult(result, dataType, result.IsInMemoryRange ? null : result.Address);
        }
        protected CompileResult CreateResult(eErrorType errorType)
        {
            return CompileResult.GetErrorResult(errorType);
        }

        /// <summary>
        /// if the supplied <paramref name="arg">argument</paramref> contains an Excel error
        /// an <see cref="ExcelErrorValueException"/> with that errorcode will be thrown
        /// </summary>
        /// <param name="arg"></param>
        /// <param name="err">If the cell contains an error the error will be assigned to this variable</param>
        protected void CheckForAndHandleExcelError(FunctionArgument arg, out ExcelErrorValue err)
        {
            err = default;
            if (arg.ValueIsExcelError)
            {
                err = arg.ValueAsExcelErrorValue;
            }
        }

        /// <summary>
        /// If the supplied <paramref name="cell"/> contains an Excel error
        /// an <see cref="ExcelErrorValueException"/> with that errorcode will be thrown
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="err">If the cell contains an error the error will be assigned to this variable</param>
        protected void CheckForAndHandleExcelError(ICellInfo cell, out ExcelErrorValue err)
        {
            err = default;
            if (cell.IsExcelError)
            {
                err = ExcelErrorValue.Parse(cell.Value.ToString());
            }
        }

        protected CompileResult GetResultByObject(object result)
        {
            if (IsNumeric(result))
            {
                return CreateResult(result, DataType.Decimal);
            }
            if (result is string)
            {
                return CreateResult(result, DataType.String);
            }
            if (ExcelErrorValue.Values.IsErrorValue(result))
            {
                return CreateResult(result, DataType.ExcelError);
            }
            if (result == null)
            {
                return CompileResult.Empty;
            }
            return CreateResult(result, DataType.ExcelRange);
        }
        /// <summary>
        /// If the function returns a different value with the same parameters.
        /// </summary>
        public virtual bool IsVolatile
        {
            get
            {
                return false;
            }
        }
        /// <summary>
        /// If the function returns a range reference
        /// </summary>
        public virtual bool ReturnsReference
        {
            get
            {
                return false;
            }
        }
        /// <summary>
        /// Provides information about the functions parameters.
        /// </summary>
        public virtual ExcelFunctionParametersInfo ParametersInfo
        {
            get;
        } = ExcelFunctionParametersInfo.Default;

        /// <summary>
        /// Information of individual arguments of the function used internally by the formula parser .
        /// </summary>
        /// <param name="argumentIndex">The argument index</param>
        /// <returns>Function argument information</returns>
        public virtual string NamespacePrefix
        {
            get
            {
                return "";
            }
        }

    }
}
