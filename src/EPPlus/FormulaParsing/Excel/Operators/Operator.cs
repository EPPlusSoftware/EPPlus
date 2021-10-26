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

namespace OfficeOpenXml.FormulaParsing.Excel.Operators
{
    public class Operator : IOperator
    {
        private const int PrecedencePercent = 2;
        private const int PrecedenceExp = 4;
        private const int PrecedenceMultiplyDevide = 6;
        private const int PrecedenceIntegerDivision = 8;
        private const int PrecedenceModulus = 10;
        private const int PrecedenceAddSubtract = 12;
        private const int PrecedenceConcat = 15;
        private const int PrecedenceComparison = 25;

        private Operator() { }

        private Operator(Operators @operator, int precedence, Func<CompileResult, CompileResult, CompileResult> implementation)
        {
            _implementation = implementation;
            _precedence = precedence;
            _operator = @operator;
        }

        private readonly Func<CompileResult, CompileResult, CompileResult> _implementation;
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

        public CompileResult Apply(CompileResult left, CompileResult right)
        {
            if (left.Result is ExcelErrorValue)
            {
                return new CompileResult(left.Result, DataType.ExcelError);
            }
            else if (right.Result is ExcelErrorValue)
            {
                return new CompileResult(right.Result, DataType.ExcelError);
            }
            return _implementation(left, right);
        }

        public override string ToString()
        {
            return "Operator: " + _operator;
        }

        private static bool CanDoNumericOperation(CompileResult l, CompileResult r)
        {
            return (l.IsNumeric || l.IsNumericString || l.IsDateString || l.Result is ExcelDataProvider.IRangeInfo) &&
                (r.IsNumeric || r.IsNumericString || r.IsDateString || r.Result is ExcelDataProvider.IRangeInfo);
        }

        private static IOperator _plus;
        public static IOperator Plus
        {
            get
            {
                return _plus ?? (_plus = new Operator(Operators.Plus, PrecedenceAddSubtract, (l, r) =>
                {
                    l = l == null || l.Result == null ? new CompileResult(0, DataType.Integer) : l;
                    r = r == null || r.Result == null ? new CompileResult(0, DataType.Integer) : r;
                    ExcelErrorValue errorVal;
                    if (EitherIsError(l, r, out errorVal))
                    {
                        return new CompileResult(errorVal);
                    }
                    if (l.DataType == DataType.Integer && r.DataType == DataType.Integer)
                    {
                        return new CompileResult(l.ResultNumeric + r.ResultNumeric, DataType.Integer);
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
                return _minus ?? (_minus = new Operator(Operators.Minus, PrecedenceAddSubtract, (l, r) =>
                {
                    l = l == null || l.Result == null ? new CompileResult(0, DataType.Integer) : l;
                    r = r == null || r.Result == null ? new CompileResult(0, DataType.Integer) : r;
                    if (l.DataType == DataType.Integer && r.DataType == DataType.Integer)
                    {
                        return new CompileResult(l.ResultNumeric - r.ResultNumeric, DataType.Integer);
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
                return _multiply ?? (_multiply = new Operator(Operators.Multiply, PrecedenceMultiplyDevide, (l, r) =>
                {
                    l = l ?? new CompileResult(0, DataType.Integer);
                    r = r ?? new CompileResult(0, DataType.Integer);
                    if (l.DataType == DataType.Integer && r.DataType == DataType.Integer)
                    {
                        return new CompileResult(l.ResultNumeric*r.ResultNumeric, DataType.Integer);
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
                return _divide ?? (_divide = new Operator(Operators.Divide, PrecedenceMultiplyDevide, (l, r) =>
                {
                    if (!(l.IsNumeric || l.IsNumericString || l.IsDateString || l.Result is ExcelDataProvider.IRangeInfo) ||
                        !(r.IsNumeric || r.IsNumericString || r.IsDateString || r.Result is ExcelDataProvider.IRangeInfo))
                    {
                        return new CompileResult(eErrorType.Value);
                    }
                    var left = l.ResultNumeric;
                    var right = r.ResultNumeric;
                    if (Math.Abs(right - 0d) < double.Epsilon)
                    {
                        return new CompileResult(eErrorType.Div0);
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
                return new Operator(Operators.Exponentiation, PrecedenceExp, (l, r) =>
                    {
                        if (l == null && r == null)
                        {
                            return new CompileResult(eErrorType.Value);
                        }
                        l = l ?? new CompileResult(0, DataType.Integer);
                        r = r ?? new CompileResult(0, DataType.Integer);
                        if (CanDoNumericOperation(l, r))
                        {
                            return new CompileResult(Math.Pow(l.ResultNumeric, r.ResultNumeric), DataType.Decimal);
                        }
                        return new CompileResult(0d, DataType.Decimal);
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
                return new Operator(Operators.Concat, PrecedenceConcat, (l, r) =>
                    {
                        l = l ?? new CompileResult(string.Empty, DataType.String);
                        r = r ?? new CompileResult(string.Empty, DataType.String);
                        var lStr = l.Result != null ? CompileResultToString(l) : string.Empty;
                        var rStr = r.Result != null ? CompileResultToString(r) : string.Empty;
                        return new CompileResult(string.Concat(lStr, rStr), DataType.String);
                    });
            }
        }

        private static IOperator _greaterThan;
        public static IOperator GreaterThan
        {
            get
            {
                return _greaterThan ??
                       (_greaterThan =
                           new Operator(Operators.GreaterThan, PrecedenceComparison,
                               (l, r) => Compare(l, r, (compRes) => compRes > 0)));
            }
        }

        private static IOperator _eq;
        public static IOperator Eq
        {
            get
            {
                return _eq ??
                       (_eq =
                           new Operator(Operators.Equals, PrecedenceComparison,
                               (l, r) => Compare(l, r, (compRes) => compRes == 0)));
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
                               (l, r) => Compare(l, r, (compRes) => compRes != 0)));
            }
        }

        private static IOperator _greaterThanOrEqual;
        public static IOperator GreaterThanOrEqual
        {
            get
            {
                return _greaterThanOrEqual ??
                       (_greaterThanOrEqual =
                           new Operator(Operators.GreaterThanOrEqual, PrecedenceComparison,
                               (l, r) => Compare(l, r, (compRes) => compRes >= 0)));
            }
        }

        private static IOperator _lessThan;
        public static IOperator LessThan
        {
            get
            {
                return _lessThan ??
                       (_lessThan =
                           new Operator(Operators.LessThan, PrecedenceComparison,
                               (l, r) => Compare(l, r, (compRes) => compRes < 0)));
            }
        }

        public static IOperator LessThanOrEqual
        {
            get
            {
                //return new Operator(Operators.LessThanOrEqual, PrecedenceComparison, (l, r) => new CompileResult(Compare(l, r) <= 0, DataType.Boolean));
                return new Operator(Operators.LessThanOrEqual, PrecedenceComparison, (l, r) => Compare(l, r, (compRes) => compRes <= 0));
            }
        }

        private static IOperator _percent;
        public static IOperator Percent
        {
            get
            {
                if (_percent == null)
                {
                    _percent = new Operator(Operators.Percent, PrecedencePercent, (l, r) =>
                        {
                            l = l ?? new CompileResult(0, DataType.Integer);
                            r = r ?? new CompileResult(0, DataType.Integer);
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
