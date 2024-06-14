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
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
using OfficeOpenXml.Table.PivotTable;

namespace OfficeOpenXml.FormulaParsing.FormulaExpressions
{
    internal static class CompileResultFactory
    {
        public static CompileResult Create(object obj)
        {
            if ((obj is INameInfo))
            {
                obj = ((INameInfo)obj).Value;
            }
            var dt =  GetDataType(ref obj);
            if(dt == DataType.ExcelRange)
            {
                var r = obj as IRangeInfo;
                return r.IsInMemoryRange ? new CompileResult(r, dt) : new AddressCompileResult(r, dt, r.Address);
            }
            return new CompileResult(obj, dt);
        }
        public static CompileResult CreateDynamicArray(object obj, FormulaRangeAddress address=null)
        {
            if ((obj is INameInfo))
            {
                obj = ((INameInfo)obj).Value;
            }
            var dt = GetDataType(ref obj);
            return new DynamicArrayCompileResult(obj, dt);
        }
        public static CompileResult Create(object obj, FormulaRangeAddress address)
        {
            bool isHidden = false;
            if ((obj is INameInfo ni))
            {
                obj = ni.Value;
            }

            var dt = GetDataType(ref obj);
            return new AddressCompileResult(obj, dt, address) { IsHiddenCell = isHidden };
        }
        private static DataType GetDataType(ref object obj)
        {
            if (obj == null) return DataType.Empty;
            var t = obj.GetType();
            var tc = Type.GetTypeCode(t);
            switch (tc)
            {
                case TypeCode.String:
                    return DataType.String;
                case TypeCode.Double:
                case TypeCode.Decimal:
                case TypeCode.Single:
                    return DataType.Decimal;
                case TypeCode.Int16:
                case TypeCode.Int32:
                case TypeCode.Int64:
                    return DataType.Integer;
                case TypeCode.Boolean:
                    return DataType.Boolean;
                case TypeCode.DateTime:
                    obj = ((DateTime)obj).ToOADate();
                    return DataType.Date;
                default:
                    if (obj is ExcelErrorValue)
                    {
                        return DataType.ExcelError;
                    }
                    else if (obj is IRangeInfo)
                    {
                        return DataType.ExcelRange;
                    }
                    throw new ArgumentException("Non supported type " + t.FullName);
            }
        }
    }
}
