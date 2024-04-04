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

namespace OfficeOpenXml.FormulaParsing.Excel.Operators
{
    public class OperatorsDict : Dictionary<string, IOperator>
    {
        public OperatorsDict()
        {
            Add("+", Operator.Plus);
            Add("-", Operator.Minus);
            Add("*", Operator.Multiply);
            Add("/", Operator.Divide);
            Add("^", Operator.Exp);
            Add("=", Operator.Eq);
            Add(">", Operator.GreaterThan);
            Add(">=", Operator.GreaterThanOrEqual);
            Add("<", Operator.LessThan);
            Add("<=", Operator.LessThanOrEqual);
            Add("<>", Operator.NotEqualsTo);
            Add("&", Operator.Concat);
            Add(":", Operator.Colon);
            Add(Operator.IntersectIndicator, Operator.Intersect);
        }

        private static IDictionary<string, IOperator> _instance;

        /// <summary>
        /// Instance of the OperatorsDict
        /// </summary>
        public static IDictionary<string, IOperator> Instance
        {
            get 
            {
                if (_instance == null)
                {
                    _instance = new OperatorsDict();
                }
                return _instance;
            }
        }
    }
    public class OperatorsEnumDict : Dictionary<Operators, IOperator>
    {
        public OperatorsEnumDict()
        {
            Add(Operators.Plus, Operator.Plus);
            Add(Operators.Minus, Operator.Minus);
            Add(Operators.Multiply, Operator.Multiply);
            Add(Operators.Divide, Operator.Divide);
            Add(Operators.Exponentiation, Operator.Exp);
            Add(Operators.Equals, Operator.Eq);
            Add(Operators.GreaterThan, Operator.GreaterThan);
            Add(Operators.GreaterThanOrEqual, Operator.GreaterThanOrEqual);
            Add(Operators.LessThan, Operator.LessThan);
            Add(Operators.LessThanOrEqual, Operator.LessThanOrEqual);
            Add(Operators.NotEqualTo, Operator.NotEqualsTo);
            Add(Operators.Concat, Operator.Concat);
            Add(Operators.Colon, Operator.Colon);
            Add(Operators.Intersect, Operator.Intersect);
        }

        private static IDictionary<Operators, IOperator> _instance;

        /// <summary>
        /// Instance of the OperatorsDict
        /// </summary>
        public static IDictionary<Operators, IOperator> Instance
        {
            get
            {
                if (_instance == null)
                {
                    _instance = new OperatorsEnumDict();
                }
                return _instance;
            }
        }
    }
}
