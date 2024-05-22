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
        public VariableExpression(string variableName, VariableFunctionExpression expression, bool isDeclaration)
        {
            Name = variableName;
            expression.DeclareVariable(variableName);
            _variableFunctionExpression = expression;
            IsDeclaration = isDeclaration;
        }

        private readonly VariableFunctionExpression _variableFunctionExpression;
        private bool _negate = false;


        internal override ExpressionType ExpressionType => ExpressionType.Variable;

        public bool IsDeclaration
        {
            get; private set;
        }

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
            return _negate ? Value.Negate() : Value;
        }

        public override Expression Negate()
        {
            _negate = !_negate;
            return this;
        }
    }
}
