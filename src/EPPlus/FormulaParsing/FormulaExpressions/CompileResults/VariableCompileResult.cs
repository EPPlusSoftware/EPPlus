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
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
using System.Diagnostics;

namespace OfficeOpenXml.FormulaParsing.FormulaExpressions.CompileResults
{
    /// <summary>
    /// CompileResult that represents a variable
    /// </summary>
    [DebuggerDisplay("Variable: {VariableName}, Result: {ResultValue}, DataType: {DataType}")]
    public class VariableCompileResult : AddressCompileResult
    {
        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="variableName">name of the variable</param>
        /// <param name="result">The compiled result</param>
        /// <param name="dataType">Data type of the compiled result</param>
        /// <param name="address">Address of the calculated cell</param>
        public VariableCompileResult(string variableName, object result, DataType dataType, FormulaRangeAddress address) : base(result, dataType, address)
        {
            VariableName = variableName;
        }

        /// <summary>
        /// Name of the variable
        /// </summary>
        public string VariableName { get; }

        /// <summary>
        /// Overrides <see cref="CompileResult.IsVariableResult"/>. For this class always true.
        /// </summary>
        public override bool IsVariableResult => true;
    }
}
