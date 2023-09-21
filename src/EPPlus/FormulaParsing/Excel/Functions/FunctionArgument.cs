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
using OfficeOpenXml.FormulaParsing.FormulaExpressions;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
using OfficeOpenXml.FormulaParsing.Ranges;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions
{
    public class FunctionArgument
    {
        public FunctionArgument(CompileResult result)
        {            
            _result= result;
            if (result.IsHiddenCell) //TODO: check if we can remove this and check result instead.
            {
                SetExcelStateFlag(Excel.ExcelCellState.HiddenCell);
            }
        }
        internal FunctionArgument(object val)
        {
            _result = CompileResultFactory.Create(val);
        }

        public FunctionArgument(object val, DataType dataType)            
        {
            _result = new CompileResult(val, dataType);
        }

        private ExcelCellState _excelCellState;

        public void SetExcelStateFlag(ExcelCellState state)
        {
            _excelCellState |= state;
        }

        public bool ExcelStateFlagIsSet(ExcelCellState state)
        {
            return (_excelCellState & state) != 0;
        }

        /// <summary>
        /// Always a IRangeInfo, even if the cell is a single cell. 
        /// <seealso cref="ValueAsRangeInfo"/>
        /// </summary>
        /// <param name="context">The parsing context</param>
        /// <returns>A <see cref="RangeInfo"/> if the argument is a range otherwise null</returns>
        public IRangeInfo GetAsRangeInfo(ParsingContext context)
        {
            if(Value is IRangeInfo ri)
            {
                return ri;
            }
            else
            {
                if(Address==null)
                {
                    return null;
                }
                else
                {
                    return new RangeInfo(Address);
                }
            }
        }

        CompileResult _result;
        public object Value { get => _result.Result; }

        public DataType DataType { get => _result.DataType; }
     
        public FormulaRangeAddress Address { get => _result.Address; } 
        public bool IsExcelRange
        {
            get => _result.DataType == DataType.ExcelRange;
        }
        public bool IsExcelRangeOrSingleCell
        {
            get => _result.DataType == DataType.ExcelRange || _result.Address != null;
        }

        public bool IsEnumerableOfFuncArgs
        {
            get { return _result.Result != null && _result.Result is IEnumerable<FunctionArgument>; }
        }

        public IEnumerable<FunctionArgument> ValueAsEnumerableOfFuncArgs
        {
            get { return _result.Result as IEnumerable<FunctionArgument>; }
        }

        public bool ValueIsExcelError
        {
            get { return ExcelErrorValue.Values.IsErrorValue(Value); }
        }

        public ExcelErrorValue ValueAsExcelErrorValue
        {
            get { return ExcelErrorValue.Parse(_result.Result.ToString()); }
        }

        /// <summary>
        /// If <see cref="Value"/> is an instance of <see cref="IRangeInfo"/> this will return a typed instance. If not null will be returned.
        /// <seealso cref="GetAsRangeInfo(ParsingContext)"/>
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
                        return new RangeInfo(_result.Address);
                    }
                    return null;
                }
            }
        }
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

        public string ValueFirstString
        {
            get
            {
                var v = ValueFirst;
                if (v == null) return default;
                return ValueFirst.ToString();
            }
        }

    }
}
