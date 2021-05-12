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
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions
{
    public class FunctionArgument
    {
        public FunctionArgument(object val)
        {
            Value = val;
            DataType = DataType.Unknown;
        }

        public FunctionArgument(object val, DataType dataType)
            :this(val)
        {
            DataType = dataType;
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

        public object Value { get; private set; }

        public DataType DataType { get; }

        public Type Type
        {
            get { return Value != null ? Value.GetType() : null; }
        }

        public int ExcelAddressReferenceId { get; set; }

        public bool IsExcelRange
        {
            get { return Value != null && Value is EpplusExcelDataProvider.IRangeInfo; }
        }

        public bool IsEnumerableOfFuncArgs
        {
            get { return Value != null && Value is IEnumerable<FunctionArgument>; }
        }

        public IEnumerable<FunctionArgument> ValueAsEnumerableOfFuncArgs
        {
            get { return Value as IEnumerable<FunctionArgument>; }
        }

        public bool ValueIsExcelError
        {
            get { return ExcelErrorValue.Values.IsErrorValue(Value); }
        }

        public ExcelErrorValue ValueAsExcelErrorValue
        {
            get { return ExcelErrorValue.Parse(Value.ToString()); }
        }

        public EpplusExcelDataProvider.IRangeInfo ValueAsRangeInfo
        {
            get { return Value as EpplusExcelDataProvider.IRangeInfo; }
        }
        public object ValueFirst
        {
            get
            {
                if (Value is ExcelDataProvider.INameInfo)
                {
                    Value = ((ExcelDataProvider.INameInfo)Value).Value;
                }
                var v = Value as ExcelDataProvider.IRangeInfo;
                if (v==null)
                {
                    return Value;
                }
                else
                {
                    return v.GetValue(v.Address._fromRow, v.Address._fromCol);
                }
            }
        }

        public string ValueFirstString
        {
            get
            {
                var v = ValueFirst;
                if (v == null) return default(string);
                return ValueFirst.ToString();
            }
        }

    }
}
