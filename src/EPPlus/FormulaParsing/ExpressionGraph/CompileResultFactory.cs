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

namespace OfficeOpenXml.FormulaParsing.ExpressionGraph
{
    public class CompileResultFactory
    {
        public virtual CompileResult Create(object obj)
        {
            return Create(obj, 0);
        }

        public virtual CompileResult Create(object obj, int excelAddressReferenceId)
        {
            if ((obj is INameInfo))
            {
                obj = ((INameInfo)obj).Value;
            }
            if (obj is IRangeInfo)
            {
                obj = ((IRangeInfo)obj).GetOffset(0, 0);
            }
            if (obj == null) return new CompileResult(null, DataType.Empty);
            if (obj.GetType().Equals(typeof(string)))
            {
                return new CompileResult(obj, DataType.String, excelAddressReferenceId);
            }
            if (obj.GetType().Equals(typeof(double)) || obj is decimal || obj is float)
            {
                return new CompileResult(obj, DataType.Decimal, excelAddressReferenceId);
            }
            if (obj.GetType().Equals(typeof(int)) || obj is long || obj is short)
            {
                return new CompileResult(obj, DataType.Integer, excelAddressReferenceId);
            }
            if (obj.GetType().Equals(typeof(bool)))
            {
                return new CompileResult(obj, DataType.Boolean, excelAddressReferenceId);
            }
            if (obj.GetType().Equals(typeof (ExcelErrorValue)))
            {
                return new CompileResult(obj, DataType.ExcelError, excelAddressReferenceId);
            }
            if (obj.GetType().Equals(typeof(System.DateTime)))
            {
                return new CompileResult(((System.DateTime)obj).ToOADate(), DataType.Date, excelAddressReferenceId);
            }
            throw new ArgumentException("Non supported type " + obj.GetType().FullName);
        }
    }
}
