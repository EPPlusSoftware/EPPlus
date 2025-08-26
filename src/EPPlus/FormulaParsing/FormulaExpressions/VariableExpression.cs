/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  08/20/2025         EPPlus Software AB       Initial release EPPlus 8.2
 *************************************************************************************************/
using OfficeOpenXml.FormulaParsing.FormulaExpressions.CompileResults;
using OfficeOpenXml.FormulaParsing.FormulaExpressions.VariableStorage;
using System.Diagnostics;

namespace OfficeOpenXml.FormulaParsing.FormulaExpressions
{
    [DebuggerDisplay("VariableExpression - Name: {Name}, Value: {Value}")]
    internal class VariableExpression : Expression
    {
        public VariableExpression(string variableName, VariableFunctionExpression expression, bool isDeclaration)
        {
            Name = variableName;
            expression.DeclareVariable(variableName);
            _variableFunctionExpression = expression;
            IsDeclaration = isDeclaration;
        }

        public VariableExpression(string variableName, VariableStorageScope scope, bool isDeclaration)
        {
            Name = variableName;
            _scope = scope;
            IsDeclaration = isDeclaration;
        }

        private readonly VariableFunctionExpression _variableFunctionExpression;
        private readonly VariableStorageScope _scope;
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
                return GetValue(out bool hasValue);
            }
        }

        private CompileResult GetValue(out bool hasValue)
        {
            hasValue = false;
            if (_scope != null)
            {
                var v = _scope.GetVariableValue(Name);
                hasValue = v.DataType != DataType.Empty && v.ResultValue != null;
                return new VariableCompileResult(Name, v.ResultValue, v.DataType, v.Address);
            }
            else if (_variableFunctionExpression != null)
            {
                var v =  _variableFunctionExpression.GetVariableValue(Name);
                hasValue = v.DataType != DataType.Empty && v.ResultValue != null;
                return new VariableCompileResult(Name, v.ResultValue, v.DataType, v.Address);
            }
            return new VariableCompileResult(Name, null, DataType.Empty, null);
        }

        internal void SetValue(string variableName, CompileResult value)
        {
            if(_variableFunctionExpression != null)
            {
                _variableFunctionExpression.AddVariableValue(variableName, value);
            }
        }



        internal string Name { get; private set; }

        public override CompileResult Compile()
        {
            if ((Status & ExpressionStatus.IsLambdaVariableDeclaration) == ExpressionStatus.IsLambdaVariableDeclaration)
            {
                var val = GetValue(out bool hasValue);
                if(!hasValue)
                {
                    return new CompileResult(Name, DataType.LambdaVariableDeclaration);
                }
               
            }
            return _negate ? Value.Negate() : Value;
        }

        public override Expression Negate()
        {
            _negate = !_negate;
            return this;
        }
    }
}
