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
using OfficeOpenXml.FormulaParsing.Exceptions;
using OfficeOpenXml.Utils;
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
using static OfficeOpenXml.ExcelAddressBase;
using System.Diagnostics;
using OfficeOpenXml.FormulaParsing.Ranges;

namespace OfficeOpenXml.FormulaParsing.Excel.Operators
{
    /// <summary>
    /// Implementation of operators in formula calculation.
    /// </summary>
    [DebuggerDisplay("Operator: {GetOperator()}")]
    public class Operator : IOperator
    {
        private const int PrecedenceColon = 0;
        private const int PrecedenceIntersect = 1;
        private const int PrecedencePercent = 3;
        private const int PrecedenceExp = 4;
        private const int PrecedenceMultiplyDivide = 6;
        private const int PrecedenceIntegerDivision = 8;
        private const int PrecedenceModulus = 10;
        private const int PrecedenceAddSubtract = 12;
        private const int PrecedenceConcat = 15;
        private const int PrecedenceComparison = 25;

        private Operator() { }

        private Operator(Operators @operator, int precedence, Func<CompileResult, CompileResult, ParsingContext, CompileResult> implementation)
        {
            _implementation = implementation;
            _precedence = precedence;
            _operator = @operator;
        }


        private readonly Func<CompileResult, CompileResult, ParsingContext, CompileResult> _implementation;
        private readonly int _precedence;
        private readonly Operators _operator;

        int IOperator.Precedence
        {
            get { return _precedence; }
        }

        Operators IOperator.Operator
        {
            get { return _operator; }
        }

        internal Operators GetOperator()
        {
            return _operator;
        }


        public CompileResult Apply(CompileResult left, CompileResult right, ParsingContext ctx)
        {
            if (left.Result is ExcelErrorValue)
            {
                return new CompileResult(left.Result, DataType.ExcelError);
            }
            else if (right.Result is ExcelErrorValue)
            {
                return new CompileResult(right.Result, DataType.ExcelError);
            }
            return _implementation(left, right, ctx);
        }

        private static bool CanDoNumericOperation(CompileResult l, CompileResult r)
        {
            return (l.IsNumeric || l.IsNumericString || l.IsPercentageString || l.IsDateString || l.Result is IRangeInfo) &&
                (r.IsNumeric || r.IsNumericString || r.IsPercentageString || r.IsDateString || r.Result is IRangeInfo);
        }

        private static IOperator _plus;
        public static IOperator Plus
        {
            get
            {
                return _plus ?? (_plus = new Operator(Operators.Plus, PrecedenceAddSubtract, (l, r, ctx) =>
                {
                    l = l == null || l.Result == null ? CompileResult.ZeroInt : l;
                    r = r == null || r.Result == null ? CompileResult.ZeroInt : r;
                    ExcelErrorValue errorVal;
                    if (EitherIsError(l, r, out errorVal))
                    {
                        return new CompileResult(errorVal);
                    }
                    if (l.DataType == DataType.Integer && r.DataType == DataType.Integer)
                    {
                        return new CompileResult(l.ResultNumeric + r.ResultNumeric, DataType.Integer);
                    }
                    if(l.DataType == DataType.ExcelRange || r.DataType == DataType.ExcelRange)
                    {
                        return RangeOperationsOperator.Apply(l, r, Operators.Plus, ctx);
                    }
                    else if (CanDoNumericOperation(l, r))
                    {
                        return new CompileResult(l.ResultNumeric + r.ResultNumeric, DataType.Decimal);
                    }
                    return new CompileResult(eErrorType.Value);
                }));
            }
        }

        private static IOperator _minus;
        public static IOperator Minus
        {
            get
            {
                return _minus ?? (_minus = new Operator(Operators.Minus, PrecedenceAddSubtract, (l, r, ctx) =>
                {
                    l = l == null || l.Result == null ? CompileResult.ZeroInt : l;
                    r = r == null || r.Result == null ? CompileResult.ZeroInt : r;
                    if (l.DataType == DataType.Integer && r.DataType == DataType.Integer)
                    {
                        return new CompileResult(l.ResultNumeric - r.ResultNumeric, DataType.Integer);
                    }
                    if (l.DataType == DataType.ExcelRange || r.DataType == DataType.ExcelRange)
                    {
                        return RangeOperationsOperator.Apply(l, r, Operators.Minus, ctx);
                    }
                    else if (CanDoNumericOperation(l, r))
                    {
                        return new CompileResult(l.ResultNumeric - r.ResultNumeric, DataType.Decimal);
                    }

                    return new CompileResult(eErrorType.Value);
                }));
            }
        }

        private static IOperator _multiply;
        public static IOperator Multiply
        {
            get
            {
                return _multiply ?? (_multiply = new Operator(Operators.Multiply, PrecedenceMultiplyDivide, (l, r, ctx) =>
                {
                    l = l ?? CompileResult.ZeroInt;
                    r = r ?? CompileResult.ZeroInt;
                    if (l.DataType == DataType.Integer && r.DataType == DataType.Integer)
                    {
                        return new CompileResult(l.ResultNumeric*r.ResultNumeric, DataType.Integer);
                    }
                    if (l.DataType == DataType.ExcelRange || r.DataType == DataType.ExcelRange)
                    {
                        return RangeOperationsOperator.Apply(l, r, Operators.Multiply, ctx);
                    }
                    else if (CanDoNumericOperation(l, r))
                    {
                        return new CompileResult(l.ResultNumeric*r.ResultNumeric, DataType.Decimal);
                    }
                    return new CompileResult(eErrorType.Value);
                }));
            }
        }

        private static IOperator _divide;
        public static IOperator Divide
        {
            get
            {
                return _divide ?? (_divide = new Operator(Operators.Divide, PrecedenceMultiplyDivide, (l, r, ctx) =>
                {
                    if (!(l.IsNumeric || l.IsNumericString || l.IsDateString || l.Result is IRangeInfo) ||
                        !(r.IsNumeric || r.IsNumericString || r.IsDateString || r.Result is IRangeInfo))
                    {
                        return new CompileResult(eErrorType.Value);
                    }
                    var left = l.ResultNumeric;
                    var right = r.ResultNumeric;
                    if (Math.Abs(right - 0d) < double.Epsilon)
                    {
                        return new CompileResult(eErrorType.Div0);
                    }
                    if (l.DataType == DataType.ExcelRange || r.DataType == DataType.ExcelRange)
                    {
                        return RangeOperationsOperator.Apply(l, r, Operators.Divide, ctx);
                    }
                    else if (CanDoNumericOperation(l, r))
                    {
                        return new CompileResult(left/right, DataType.Decimal);
                    }
                    return new CompileResult(eErrorType.Value);
                }));
            }
        }

        public static IOperator Exp
        {
            get
            {
                return new Operator(Operators.Exponentiation, PrecedenceExp, (l, r, ctx) =>
                    {
                        if (l == null && r == null)
                        {
                            return new CompileResult(eErrorType.Value);
                        }
                        l = l ?? CompileResult.ZeroInt;
                        r = r ?? CompileResult.ZeroInt;
                        if (l.DataType == DataType.ExcelRange || r.DataType == DataType.ExcelRange)
                        {
                            return RangeOperationsOperator.Apply(l, r, Operators.Exponentiation, ctx);
                        }
                        if (CanDoNumericOperation(l, r))
                        {
                            return new CompileResult(Math.Pow(l.ResultNumeric, r.ResultNumeric), DataType.Decimal);
                        }
                        return CompileResult.ZeroDecimal;
                    });
            }
        }

        private static string CompileResultToString(CompileResult c)
        {
            if(c != null && c.IsNumeric)
            {
                if(c.ResultNumeric is double d)
                {
                    return d.ToString("G15");
                }
            }
            return c.ResultValue.ToString();
        }

        public static IOperator Concat
        {
            get
            {
                return new Operator(Operators.Concat, PrecedenceConcat, (l, r, ctx) =>
                    {
                        l = l ?? new CompileResult(string.Empty, DataType.String);
                        r = r ?? new CompileResult(string.Empty, DataType.String);
                        if (l.DataType == DataType.ExcelRange || r.DataType == DataType.ExcelRange)
                        {
                            return RangeOperationsOperator.Apply(l, r, Operators.Concat, ctx);
                        }
                        var lStr = l.Result != null ? CompileResultToString(l) : string.Empty;
                        var rStr = r.Result != null ? CompileResultToString(r) : string.Empty;
                        return new CompileResult(string.Concat(lStr, rStr), DataType.String);
                    });
            }
        }

        static IOperator _colon=null;
        public static IOperator Colon
        {
            get
            {
                if (_colon == null)
                {
                    _colon = new Operator(Operators.Colon, PrecedenceColon, (l, r, ctx) =>
                      {
                          FormulaRangeAddress result;
                          if (l.Address != null)
                          {
                              result = l.Address;
                              if (result.Address.WorksheetIx < -1)
                              {
                                  result.WorksheetIx = (short)ctx.CurrentCell.WorksheetIx;
                              }
                          }
                          else
                          {
                              return new AddressCompileResult(eErrorType.Value);
                          }
                          //else
                          //{
                          //    if (l.Result is IRangeInfo lri)
                          //    {
                          //        result = new FormulaRangeAddress(ctx);
                          //        result.WorksheetIx = lri.Address.WorksheetIx < -1 ? (short)ctx.CurrentCell.WorksheetIx : lri.Address.WorksheetIx;
                          //        result.FromRow = lri.Address.FromRow;
                          //        result.FromCol = lri.Address.FromCol;
                          //        result.ToRow = lri.Address.ToRow;
                          //        result.ToCol = lri.Address.ToCol;
                          //    }
                          //}
                          
                          if (r.Address != null)
                          {
                              if (result.WorksheetIx != r.Address.WorksheetIx && r.Address.WorksheetIx != short.MinValue)
                              {
                                  result.WorksheetIx = -1;
                              }
                              else
                              {
                                  result.FromRow = result.FromRow < r.Address.FromRow ? result.FromRow : r.Address.FromRow;
                                  result.FromCol = result.FromCol < r.Address.FromCol ? result.FromCol : r.Address.FromCol;
                                  result.ToRow = result.ToRow > r.Address.ToRow ? result.ToRow : r.Address.ToRow;   
                                  result.ToCol = result.ToCol > r.Address.ToCol ? result.ToCol : r.Address.ToCol;
                                  
                              }
                          }
                          else
                          {
                                  return new AddressCompileResult(eErrorType.Value);
                          }
                          //else if(r.Address!=null)
                          //{
                          //    if (result.WorksheetIx != r.Address.WorksheetIx && r.Address.WorksheetIx!=short.MinValue)
                          //    {
                          //        result.WorksheetIx = -1;
                          //    }
                          //    else
                          //    {
                          //        result.FromRow = result.FromRow < r.Address.FromRow ? result.FromRow : r.Address.FromRow;
                          //        result.FromCol = result.FromCol < r.Address.FromCol ? result.FromCol : r.Address.FromCol;
                          //        result.ToRow = result.ToRow > r.Address.ToRow ? result.ToRow : r.Address.ToRow;
                          //        result.ToCol = result.ToCol > r.Address.ToCol ? result.ToCol : r.Address.ToCol;
                          //    }
                          //}

                          return new AddressCompileResult(new RangeInfo(result, ctx), DataType.ExcelRange,result);
                          throw new ExcelErrorValueException(eErrorType.Ref);
                      });
                }
                return _colon;
            }
        }
        //static IOperator _exclamation = null;
        //public static IOperator Exclamation
        //{
        //    get
        //    {
        //        if(_exclamation==null)
        //        {
        //            _exclamation = new Operator(Operators.Exclamation, PrecedenceExclamation,
        //                           (l, r, ctx) => 
        //                           {
        //                               if(l.DataType == DataType.WorksheetName && r.DataType == DataType.ExcelCellAddress)
        //                               {
        //                                   var wsName = l.Result?.ToString();
        //                                   var cellAddress = r.Result as FormulaCellAddress;
        //                                   if(cellAddress != null)
        //                                   {
        //                                       var wsIndex = ctx.Package.Workbook.Worksheets[wsName]?.PositionId;
        //                                       cellAddress.WorksheetIx = (short)(wsIndex.HasValue ? wsIndex.Value : -1);
        //                                       return new CompileResult(cellAddress, DataType.ExcelCellAddress);
        //                                   }
        //                               }
        //                               else if (l.DataType == DataType.WorksheetName && r.DataType == DataType.Enumerable)
        //                               {
        //                                   var wsName = l.Result?.ToString();
        //                                   var rangeInfo = r.Result as IRangeInfo;
        //                                   if (rangeInfo != null)
        //                                   {
        //                                       rangeInfo.Address._ws = wsName;                                               
        //                                       var wsIndex = ctx.Package.Workbook.Worksheets[wsName]?.PositionId;
        //                                       rangeInfo.RangeNew.WorksheetIx = (short)(wsIndex.HasValue ? wsIndex.Value : -1);
        //                                       return new CompileResult(rangeInfo, DataType.ExcelRange);
        //                                   }
        //                               }
        //                               return new CompileResult(eErrorType.Ref);
        //                           });
        //        }
        //        return _exclamation;
        //    }
        //}

        static IOperator _intersect = null;
        /// <summary>
        /// Intersect operator
        /// </summary>
        public static IOperator Intersect
        {
            get
            {
                if (_intersect == null)
                {
                    _intersect = new Operator(Operators.Intersect, PrecedenceIntersect,
                                   (l, r, ctx) =>
                                   {
                                       FormulaRangeAddress la, ra;
                                       if(l.Result is IRangeInfo left)
                                       {
                                           la = left.Address;
                                       }
                                       else if (l.Result is FormulaRangeAddress lfra)
                                       {
                                           la = lfra;
                                       }
                                       else
                                       {
                                           la = null;
                                       }

                                       if (r.Result is IRangeInfo right)
                                       {
                                           ra = right.Address;
                                       }
                                       else if (r.Result is FormulaRangeAddress rfra)
                                       {
                                           ra = rfra;
                                       }
                                       else
                                       {
                                           ra = null;
                                       }

                                       if (la!=null && ra!=null)
                                       {
                                           var iA = la.Intersect(ra);
                                           if (iA == null)
                                           {
                                               return new CompileResult(eErrorType.Null);
                                           }
                                           var intersectRange = ctx.ExcelDataProvider.GetRange(iA);
                                           return new CompileResult(intersectRange, DataType.ExcelRange);
                                           
                                       }
                                       return new CompileResult(eErrorType.Value);
                                   });
                }
                return _intersect;
            }
        }

        private static IOperator _greaterThan;
        /// <summary>
        /// Greater than operator
        /// </summary>
        public static IOperator GreaterThan
        {
            get
            {
                return _greaterThan ??
                       (_greaterThan =
                           new Operator(Operators.GreaterThan, PrecedenceComparison, (l, r, ctx) => 
                           {
                               if (l.DataType == DataType.ExcelRange || r.DataType == DataType.ExcelRange)
                               {
                                   return RangeOperationsOperator.Apply(l, r, Operators.GreaterThan, ctx);
                               }
                               return Compare(l, r, (compRes) => compRes > 0);
                           }));
            }
        }

        private static IOperator _eq;
        /// <summary>
        /// Equals operator
        /// </summary>
        public static IOperator Eq
        {
            get
            {
                return _eq ??
                       (_eq =
                           new Operator(Operators.Equals, PrecedenceComparison, (l, r, ctx) => 
                           {
                               if (l.DataType == DataType.ExcelRange || r.DataType == DataType.ExcelRange)
                               {
                                   return RangeOperationsOperator.Apply(l, r, Operators.Equals, ctx);
                               }
                               return Compare(l, r, (compRes) => compRes == 0); 
                               
                           }));
            }
        }

        private static IOperator _notEqualsTo;
        public static IOperator NotEqualsTo
        {
            get
            {
                return _notEqualsTo ??
                       (_notEqualsTo =
                           new Operator(Operators.NotEqualTo, PrecedenceComparison,
                               (l, r, ctx) => Compare(l, r, (compRes) => compRes != 0)));
            }
        }

        private static IOperator _greaterThanOrEqual;
        public static IOperator GreaterThanOrEqual
        {
            get
            {
                return _greaterThanOrEqual ??
                       (_greaterThanOrEqual =
                           new Operator(Operators.GreaterThanOrEqual, PrecedenceComparison, (l, r, ctx) => 
                           {
                               if (l.DataType == DataType.ExcelRange || r.DataType == DataType.ExcelRange)
                               {
                                   return RangeOperationsOperator.Apply(l, r, Operators.GreaterThanOrEqual, ctx);
                               }
                               return Compare(l, r, (compRes) => compRes >= 0); 
                           }));
            }
        }

        private static IOperator _lessThan;
        public static IOperator LessThan
        {
            get
            {
                return _lessThan ??
                       (_lessThan =
                           new Operator(Operators.LessThan, PrecedenceComparison, (l, r, ctx) =>
                           {
                               if (l.DataType == DataType.ExcelRange || r.DataType == DataType.ExcelRange)
                               {
                                   return RangeOperationsOperator.Apply(l, r, Operators.LessThan, ctx);
                               }
                               return Compare(l, r, (compRes) => compRes < 0);
                           }));
            }
        }

        private static IOperator _lessThanOrEqual;
        public static IOperator LessThanOrEqual
        {
            get
            {
                return _lessThanOrEqual ?? 
                    (_lessThanOrEqual = 
                        new Operator(Operators.LessThanOrEqual, PrecedenceComparison, (l, r, ctx) => 
                        {
                            if (l.DataType == DataType.ExcelRange || r.DataType == DataType.ExcelRange)
                            {
                                return RangeOperationsOperator.Apply(l, r, Operators.LessThanOrEqual, ctx);
                            }
                            return Compare(l, r, (compRes) => compRes <= 0); 
                        }));
            }
        }

        private static IOperator _percent;
        public static IOperator Percent
        {
            get
            {
                if (_percent == null)
                {
                    _percent = new Operator(Operators.Percent, PrecedencePercent, (l, r, ctx) =>
                        {
                            l = l ?? CompileResult.ZeroInt;
                            r = r ?? CompileResult.ZeroInt;
                            if (l.DataType == DataType.Integer && r.DataType == DataType.Integer)
                            {
                                return new CompileResult(l.ResultNumeric * r.ResultNumeric, DataType.Integer);
                            }
                            else if (CanDoNumericOperation(l, r))
                            {
                                return new CompileResult(l.ResultNumeric * r.ResultNumeric, DataType.Decimal);
                            }
                            return new CompileResult(eErrorType.Value);
                        });
                }
                return _percent;
            }
        }

        private static object GetObjFromOther(CompileResult obj, CompileResult other)
        {
            if (obj.Result == null)
            {
                if (other.DataType == DataType.String) return string.Empty;
                else return 0d;
            }
            return obj.ResultValue;
        }

        private static CompileResult Compare(CompileResult l, CompileResult r, Func<int, bool> comparison )
        {
            ExcelErrorValue errorVal;
            if (EitherIsError(l, r, out errorVal))
            {
                return new CompileResult(errorVal);
            }
            object left, right;
            left = GetObjFromOther(l, r);
            right = GetObjFromOther(r, l);
            if (ConvertUtil.IsNumericOrDate(left) && ConvertUtil.IsNumericOrDate(right))
            {
                var lnum = ConvertUtil.GetValueDouble(left);
                var rnum = ConvertUtil.GetValueDouble(right);
                if (Math.Abs(lnum - rnum) < double.Epsilon)
                {
                    return new CompileResult(comparison(0), DataType.Boolean);
                }
                var comparisonResult = lnum.CompareTo(rnum);
                return new CompileResult(comparison(comparisonResult), DataType.Boolean);
            }
            else
            {
                var comparisonResult = CompareString(left, right);
                return new CompileResult(comparison(comparisonResult), DataType.Boolean);
            }
        }

        private static int CompareString(object l, object r)
        {
            var sl = (l ?? "").ToString();
            var sr = (r ?? "").ToString();
            return string.Compare(sl, sr, StringComparison.OrdinalIgnoreCase);
        }

        private static bool  EitherIsError(CompileResult l, CompileResult r, out ExcelErrorValue errorVal)
        {
            if (l.DataType == DataType.ExcelError)
            {
                errorVal = (ExcelErrorValue) l.Result;
                return true;
            }
            if (r.DataType == DataType.ExcelError)
            {
                errorVal = (ExcelErrorValue) r.Result;
                return true;
            }
            errorVal = null;
            return false;
        }
    }
}
