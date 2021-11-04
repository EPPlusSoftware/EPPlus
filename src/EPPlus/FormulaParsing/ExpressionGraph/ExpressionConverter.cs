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

namespace OfficeOpenXml.FormulaParsing.ExpressionGraph
{
    public class ExpressionConverter : IExpressionConverter
    {
        public StringExpression ToStringExpression(Expression expression)
        {
            var result = expression.Compile();
            string toString;
            if(result.DataType == DataType.Decimal)
            {
                toString = result.ResultNumeric.ToString("G15");
            }
            else
            {
                toString = result.Result.ToString();
            }
            var newExp = new StringExpression(toString);
            newExp.Operator = expression.Operator;
            return newExp;
        }

        public Expression FromCompileResult(CompileResult compileResult)
        {
            switch (compileResult.DataType)
            {
                case DataType.Integer:
                    return compileResult.Result is string
                        ? new IntegerExpression(compileResult.Result.ToString())
                        : new IntegerExpression(Convert.ToDouble(compileResult.Result));
                case DataType.String:
                    return new StringExpression(compileResult.Result.ToString());
                case DataType.Decimal:
                    return compileResult.Result is string
                               ? new DecimalExpression(compileResult.Result.ToString())
                               : new DecimalExpression(((double) compileResult.Result));
                case DataType.Boolean:
                    return compileResult.Result is string
                               ? new BooleanExpression(compileResult.Result.ToString())
                               : new BooleanExpression((bool) compileResult.Result);
                //case DataType.Enumerable:
                //    return 
                case DataType.ExcelError:
                    //throw (new OfficeOpenXml.FormulaParsing.Exceptions.ExcelErrorValueException((ExcelErrorValue)compileResult.Result)); //Added JK
                    return compileResult.Result is string
                        ? new ExcelErrorExpression(compileResult.Result.ToString(),
                            ExcelErrorValue.Parse(compileResult.Result.ToString()))
                        : new ExcelErrorExpression((ExcelErrorValue) compileResult.Result);
                case DataType.Empty:
                   return new IntegerExpression(0); //Added JK
                case DataType.Time:
                case DataType.Date:
                    return new DecimalExpression((double)compileResult.Result);
                case DataType.Enumerable:
                    var rangeInfo = compileResult.Result as ExcelDataProvider.IRangeInfo;
                    if (rangeInfo != null)
                    {
                        return new ExcelRangeExpression(rangeInfo);
                    }
                    break;

            }
            return null;
        }

        private static IExpressionConverter _instance;
        public static IExpressionConverter Instance
        {
            get
            {
                if (_instance == null)
                {
                    _instance = new ExpressionConverter();
                }
                return _instance;
            }
        }
    }
}
