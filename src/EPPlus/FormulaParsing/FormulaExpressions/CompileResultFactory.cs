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

namespace OfficeOpenXml.FormulaParsing.FormulaExpressions
{
    internal static class CompileResultFactory
    {
        public static CompileResult Create(object obj)
        {
            if (obj is IRangeInfo)
            {
                obj = ((IRangeInfo)obj).GetOffset(0, 0);
            }
            else if ((obj is INameInfo))
            {
                obj = ((INameInfo)obj).Value;
            }
            if (obj == null) return new CompileResult(null, DataType.Empty);
            var t = obj.GetType();
            var tc = Type.GetTypeCode(t);
            switch (tc)
            {
                case TypeCode.String:
                   return new CompileResult(obj, DataType.String);
                case TypeCode.Double:
                case TypeCode.Decimal:
                case TypeCode.Single:
                    return new CompileResult(obj, DataType.Decimal);
                case TypeCode.Int16:
                case TypeCode.Int32:
                case TypeCode.Int64:
                    return new CompileResult(obj, DataType.Integer);
                case TypeCode.Boolean:
                    return new CompileResult(obj, DataType.Boolean);
                case TypeCode.DateTime:
                    return new CompileResult(((System.DateTime)obj).ToOADate(), DataType.Date);
                default:
                    if (t.Equals(typeof(ExcelErrorValue)))
                    {
                        return new CompileResult(obj, DataType.ExcelError);
                    }
                    throw new ArgumentException("Non supported type " + t.FullName);
            }
        }
        public static CompileResult Create(object obj, FormulaRangeAddress address)
        {
            bool isHidden=false;
            if (obj is IRangeInfo ri)
            {
                obj = ri.GetOffset(0, 0);
            }
            else if ((obj is INameInfo ni))
            {
                obj = ni.Value;
            }
            if (obj == null) return new AddressCompileResult(null, DataType.Empty, address);
            var t = obj.GetType();
            var tc = Type.GetTypeCode(t);
            switch (tc)
            {
                case TypeCode.String:
                    return new AddressCompileResult(obj, DataType.String, address) { IsHiddenCell = isHidden };
                case TypeCode.Double:
                case TypeCode.Decimal:
                case TypeCode.Single:
                    return new AddressCompileResult(obj, DataType.Decimal, address) { IsHiddenCell = isHidden };
                case TypeCode.Int16:
                case TypeCode.Int32:
                case TypeCode.Int64:
                        return new AddressCompileResult(obj, DataType.Integer, address) { IsHiddenCell = isHidden };
                case TypeCode.Boolean:
                    return new AddressCompileResult(obj, DataType.Boolean, address) { IsHiddenCell = isHidden };
                case TypeCode.DateTime:
                    return new AddressCompileResult(((System.DateTime)obj).ToOADate(), DataType.Date, address) { IsHiddenCell = isHidden };
                default:
                    if (t.Equals(typeof(ExcelErrorValue)) || t.IsSubclassOf(typeof(ExcelErrorValue)))
                    {
                        return new AddressCompileResult(obj, DataType.ExcelError, address);
                    }
                    throw new ArgumentException("Non supported type " + t.FullName);
            }
        }
    }
}
