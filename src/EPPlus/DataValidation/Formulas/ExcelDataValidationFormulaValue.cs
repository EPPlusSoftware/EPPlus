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
using OfficeOpenXml.DataValidation.Events;
using System;

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
        /// <param name="validationUid">Uid for the data validation</param>
        public ExcelDataValidationFormulaValue(string validationUid, string worksheetName, Action<OnFormulaChangedEventArgs> extListHandler)
            : base(validationUid, worksheetName, extListHandler)
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
            }
        }

        internal override void ResetValue()
        {
            Value = default(T);
        }

    }
}
