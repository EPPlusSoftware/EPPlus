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
using OfficeOpenXml.FormulaParsing.ExcelUtilities;
using OfficeOpenXml.FormulaParsing.Utilities;

namespace OfficeOpenXml.FormulaParsing.ExpressionGraph
{
    public class ExcelErrorExpression : Expression
    {
        ExcelErrorValue _error;
        public ExcelErrorExpression(string expression, ExcelErrorValue error)
            : base(expression)
        {
            _error = error;
        }

        public ExcelErrorExpression(ExcelErrorValue error)
            : this(error.ToString(), error)
        {
            
        }

        public override bool IsGroupedExpression
        {
            get { return false; }
        }

        public override CompileResult Compile()
        {
            return new CompileResult(_error, DataType.ExcelError);
            //if (ParentIsLookupFunction)
            //{
            //    return new CompileResult(ExpressionString, DataType.ExcelError);
            //}
            //else
            //{
            //    return CompileRangeValues();
            //}
        }
    }
}
