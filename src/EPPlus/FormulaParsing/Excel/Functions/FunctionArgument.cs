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
using OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup;
using OfficeOpenXml.FormulaParsing.FormulaExpressions;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
using OfficeOpenXml.FormulaParsing.Ranges;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions
{
    /// <summary>
    /// Represents a function argument passed to the Execute method of a <see cref="ExcelFunction"/> class.
    /// <see cref="ExcelFunction.Execute(IList{FunctionArgument}, ParsingContext)"/>
    /// </summary>
    public class FunctionArgument
    {
        internal FunctionArgument(CompileResult result)
        {            
            _result= result;
        }
        /// <summary>
        /// Constructor.
        /// </summary>
        /// <param name="val">The value of the function argument.</param>
        public FunctionArgument(object val)
        {
            _result = CompileResultFactory.Create(val);
        }

        /// <summary>
        /// Constructor.
        /// </summary>
        /// <param name="val">The value of the function argument.</param>
        /// <param name="dataType">The data type of the <paramref name="val"/>. The data type should match the values .NET data type</param>
        public FunctionArgument(object val, DataType dataType)            
        {
            _result = new CompileResult(val, dataType);
        }

        /// <summary>
        /// If the compile result has a function that handles hidden cells.
        /// </summary>
        public bool IsHiddenCell 
        {
            get
            {
                return _result.IsHiddenCell;
            }
            internal set 
            {
                _result.IsHiddenCell = value;
            }
        }
        CompileResult _result;
        /// <summary>
        /// The value of the function argument
        /// </summary>
        public object Value { get => _result.Result; }
        /// <summary>
        /// The data type of the <see cref="Value"/>.
        /// </summary>
        public DataType DataType { get => _result.DataType; }
        /// <summary>
        /// The address for function parameter 
        /// </summary>
        public FormulaRangeAddress Address { get => _result.Address; } 
        /// <summary>
        /// If the <see cref="Value"/> is a range with more than one cell.
        /// </summary>
        public bool IsExcelRange
        {
            get => _result.DataType == DataType.ExcelRange;
        }
        /// <summary>
        /// If the <see cref="Value"/> is a range.
        /// </summary>
        public bool IsExcelRangeOrSingleCell
        {
            get => _result.DataType == DataType.ExcelRange || _result.Address != null;
        }

        /// <summary>
        /// Returns true if the <see cref="Value"/> is an <see cref="ExcelErrorValue"/>
        /// </summary>
        public bool ValueIsExcelError
        {
            get { return ExcelErrorValue.Values.IsErrorValue(Value); }
        }

        /// <summary>
        /// Tries to parse <see cref="Value"/> as <see cref="ExcelErrorValue"/>
        /// </summary>
        public ExcelErrorValue ValueAsExcelErrorValue
        {
            get { return ExcelErrorValue.Parse(_result.Result.ToString()); }
        }

        /// <summary>
        /// If <see cref="Value"/> is an instance of <see cref="IRangeInfo"/> or has <see cref="Address"/> set to a valid address
        /// this property will return a <see cref="IRangeInfo"/>. If not null will be returned.
        /// </summary>
        public IRangeInfo ValueAsRangeInfo
        {
            get 
            {
                if(_result.Result is IRangeInfo ri)
                {
                    return ri;
                }
                else
                {
                    if(_result.Address!=null)
                    {
                        return _result.Address.GetAsRangeInfo();
                    }
                    return null;
                }
            }
        }
        /// <summary>
        /// If the value is a <see cref="IRangeInfo"/> the value will return the value of the first cell, otherwise the <see cref="Value"/> will be returned.
        ///
        /// </summary>
        public object ValueFirst
        {
            get
            {
                if (_result.Result is INameInfo ni)
                {
                    return ni.Value;
                }
                var v = _result.Result as IRangeInfo;
                if (v == null)
                {
                    return _result.Result;
                }
                else
                {
                    if (v.IsInMemoryRange)
                    {
                        return v.GetValue(0, 0);
                    }
                    else
                    {
                        return v.GetValue(v.Address.FromRow, v.Address.FromCol);
                    }
                }
            }
        }

        /// <summary>
        /// If the value is a <see cref="IRangeInfo"/> the value will return the value of the first cell, otherwise the <see cref="Value"/> will be returned.
        ///
        /// </summary>
        public List<object> ValueToList
        {
            get
            {
                List<object> obj = new List<object>();
                var v = _result.Result as IRangeInfo;
                if(v == null)
                {
                    obj.Add(_result.Result);
                    return obj;
                }
                for (int row = v.Address.FromRow; row <= v.Address.ToRow; row++)
                {
                    for (int col = v.Address.FromCol; col <= v.Address.ToCol; col++)
                    {
                        obj.Add(v.GetValue(row, col));
                    }
                }
                return obj;
            }
        }

        //public string ValueFirstString
        //{
        //    get
        //    {
        //        var v = ValueFirst;
        //        if (v == null) return default;
        //        return ValueFirst.ToString();
        //    }
        //}
    }
}
