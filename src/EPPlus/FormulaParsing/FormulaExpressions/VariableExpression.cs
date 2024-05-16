/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  05/14/2024         EPPlus Software AB       Initial release EPPlus 7.3
 *************************************************************************************************/
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.FormulaExpressions
{
    internal class VariableExpression : Expression
    {
        public VariableExpression(string variableName, VariableFunctionExpression expression)
        {
            Name = variableName;
            expression.AddVariableName(variableName);
            _variableFunctionExpression = expression;
        }

        private readonly VariableFunctionExpression _variableFunctionExpression;

        internal override ExpressionType ExpressionType => ExpressionType.Variable;

        internal override ExpressionStatus Status
        {
            get;
            set;
        } = ExpressionStatus.CanCompile;

        internal CompileResult Value
        {
            get
            {
                return _variableFunctionExpression.GetVariableValue(Name);
            }
        }

        internal string Name { get; private set; }

        public override CompileResult Compile()
        {
            return Value;
        }

        public override Expression Negate()
        {
            throw new NotImplementedException();
        }
    }
}
