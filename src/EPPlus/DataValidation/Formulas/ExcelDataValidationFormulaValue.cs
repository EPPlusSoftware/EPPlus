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

namespace OfficeOpenXml.DataValidation.Formulas
{

    /// <summary>
    /// This class represents a validation formula. Its value can be specified as a value of the specified datatype or as a formula.
    /// </summary>
    /// <typeparam name="T"></typeparam>
    internal abstract class ExcelDataValidationFormulaValue<T> : ExcelDataValidationFormula
    {
        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="namespaceManager">Namespacemanger of the worksheet</param>
        /// <param name="topNode">validation top node</param>
        /// <param name="formulaPath">xml path of the current formula</param>
        /// <param name="validationUid">Uid for the data validation</param>
        public ExcelDataValidationFormulaValue(string formula, string validationUid)
            : base(formula, validationUid)
        {

        }

        private T _value;
        /// <summary>
        /// Typed value
        /// </summary>
        public T Value
        {
            get
            {
                return _value;
            }
            set
            {
                State = FormulaState.Value;
                _value = value;
                _formula = GetValueAsString();
            }
        }

        internal override void ResetValue()
        {
            Value = default(T);
        }

    }
}
