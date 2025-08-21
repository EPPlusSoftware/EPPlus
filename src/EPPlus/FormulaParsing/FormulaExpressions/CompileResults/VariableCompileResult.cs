using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.FormulaExpressions.CompileResults
{
    [DebuggerDisplay("Variable: {VariableName}, Result: {ResultValue}, DataType: {DataType}")]
    public class VariableCompileResult : AddressCompileResult
    {
        public VariableCompileResult(string variableName, object result, DataType dataType, FormulaRangeAddress address) : base(result, dataType, address)
        {
            VariableName = variableName;
        }

        public string VariableName { get; }

        public override bool IsVariableResult => true;
    }
}
