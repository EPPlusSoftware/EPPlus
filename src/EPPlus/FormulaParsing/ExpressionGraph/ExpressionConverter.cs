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
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
using static OfficeOpenXml.FormulaParsing.EpplusExcelDataProvider;

namespace OfficeOpenXml.FormulaParsing.ExpressionGraph
{
    public class ExpressionConverter : IExpressionConverter
    {
        internal ExpressionConverter(ParsingContext ctx)
        {
            _ctx = ctx;
        }

        private readonly ParsingContext _ctx;

        public StringExpression ToStringExpression(Expression expression)
        {
            var result = expression.Compile();
            string toString;
            if (result.DataType == DataType.Decimal)
            {
                toString = result.ResultNumeric.ToString("G15");
            }
            else
            {
                toString = result.Result.ToString();
            }
            var newExp = new StringExpression(toString, _ctx);
            newExp.Operator = expression.Operator;
            return newExp;
        }

        public Expression FromCompileResult(CompileResult compileResult)
        {
            switch (compileResult.DataType)
            {
                case DataType.Integer:
                    return compileResult.Result is string
                        ? new IntegerExpression(compileResult.Result.ToString(), _ctx)
                        : new IntegerExpression(Convert.ToDouble(compileResult.Result), _ctx);
                case DataType.String:
                    return new StringExpression(compileResult.Result.ToString(), _ctx);
                case DataType.Decimal:
                    return compileResult.Result is string
                               ? new DecimalExpression(compileResult.Result.ToString(), _ctx)
                               : new DecimalExpression(((double)compileResult.Result), _ctx);
                case DataType.Boolean:
                    return compileResult.Result is string
                               ? new BooleanExpression(compileResult.Result.ToString(), _ctx)
                               : new BooleanExpression((bool)compileResult.Result, _ctx);
                //case DataType.Enumerable:
                //    return 
                case DataType.ExcelError:
                    //throw (new OfficeOpenXml.FormulaParsing.Exceptions.ExcelErrorValueException((ExcelErrorValue)compileResult.Result)); //Added JK
                    return compileResult.Result is string
                        ? new ExcelErrorExpression(compileResult.Result.ToString(),
                            ExcelErrorValue.Parse(compileResult.Result.ToString()), _ctx)
                        : new ExcelErrorExpression((ExcelErrorValue)compileResult.Result, _ctx);
                case DataType.Empty:
                    return new IntegerExpression(0, _ctx); //Added JK
                case DataType.Time:
                case DataType.Date:
                    return new DecimalExpression((double)compileResult.Result, _ctx);
                case DataType.Enumerable:
                    var rangeInfo = compileResult.Result as IRangeInfo;
                    if (rangeInfo != null)
                    {
                        return new ExcelRangeExpression(rangeInfo, _ctx);
                    }
                    break;
                case DataType.ExcelRange:                    
                    if(compileResult.Result is FormulaRangeAddress f)
                    {
                        if(f.ExternalReferenceIx < -1 || f.WorksheetIx==-1)
                        {
                            return new ExcelErrorExpression(ExcelErrorValue.Create(eErrorType.Ref), _ctx);
                        }
                        if (f.WorksheetIx == short.MinValue) f.WorksheetIx = _ctx.Scopes.Current.Address.WorksheetIx;
                        return new ExcelRangeExpression(_ctx.ExcelDataProvider.GetRange(_ctx.Package.Workbook.Worksheets[f.WorksheetIx]?.Name, f.FromRow, f.FromCol, f.ToRow, f.ToCol), _ctx);
                    }
                    else if (compileResult.Result is IRangeInfo ri)
                    {
                        if(ri.IsInMemoryRange)
                        {
                            return new ExcelRangeExpression(ri, _ctx);
                        }
                        return new ExcelRangeExpression(_ctx.ExcelDataProvider.GetRange(ri.Worksheet?.Name, ri.Address.FromRow, ri.Address.FromCol, ri.Address.ToRow, ri.Address.ToCol), _ctx);
                    }
                    break;
                case DataType.ExcelCellAddress:
                    if (compileResult.ResultValue is FormulaRangeAddress r)
                    {
                        if (r.ExternalReferenceIx<0 || r.WorksheetIx < 0)
                        {
                            return new ExcelErrorExpression(ExcelErrorValue.Create(eErrorType.Ref), _ctx);
                        }
                        return new ExcelRangeExpression(_ctx.ExcelDataProvider.GetRange(_ctx.Package.Workbook.Worksheets[r.WorksheetIx]?.Name, r.FromRow, r.FromCol, r.ToRow, r.ToCol), _ctx);
                    }
                    break;
            }
            return null;
        }

        //private static IExpressionConverter _instance;
        public static IExpressionConverter GetInstance(ParsingContext context)
        {
            return new ExpressionConverter(context);
        }
    }
}
