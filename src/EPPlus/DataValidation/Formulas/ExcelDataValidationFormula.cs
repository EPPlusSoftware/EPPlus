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
using OfficeOpenXml.DataValidation.Contracts;
using OfficeOpenXml.DataValidation.Events;
using OfficeOpenXml.DataValidation.Exceptions;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
using OfficeOpenXml.Utils;
using System;
using System.Linq;
using OfficeOpenXml.DataValidation.Formulas.Contracts;

namespace OfficeOpenXml.DataValidation.Formulas
{
    /// <summary>
    /// Enumeration representing the state of an <see cref="ExcelDataValidationFormulaValue{T}"/>
    /// </summary>
    internal enum FormulaState
    {
        /// <summary>
        /// Value is set
        /// </summary>
        Value,
        /// <summary>
        /// Formula is set
        /// </summary>
        Formula
    }

    /// <summary>
    /// Base class for a formula
    /// </summary>
    internal abstract class ExcelDataValidationFormula :IExcelDataValidationFormula
    {

        internal event System.EventHandler BecomesExt;

        private readonly Action<OnFormulaChangedEventArgs> _handler;

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="validationUid">id of the data validation containing this formula</param>
        public ExcelDataValidationFormula(string validationUid, string workSheetName, Action<OnFormulaChangedEventArgs> extListHandler)
        {
            Require.Argument(validationUid).IsNotNullOrEmpty("validationUid");
            _validationUid = validationUid;
            _workSheetName = workSheetName;
            _handler = extListHandler;
        }

        private string _validationUid;
        protected string _formula;
        private string _workSheetName;

        internal virtual bool HasValue { get; set; } = false;
        /// <summary>
        /// State of the validationformula, i.e. tells if value or formula is set
        /// </summary>
        protected FormulaState State
        {
            get;
            set;
        }

        private int MeasureFormulaLength(string formula)
        {
            if (string.IsNullOrEmpty(formula)) return 0;
            formula = formula.Replace("_xlfn.", string.Empty).Replace("_xlws.", string.Empty);
            return formula.Length;
        }

        /// <summary>
        /// A formula which output must match the current validation type
        /// </summary>
        public string ExcelFormula
        {
            get
            {
                return _formula;
            }
            set
            {
                if (value != null && MeasureFormulaLength(value) > 255)
                {
                    throw new DataValidationFormulaTooLongException("The length of a DataValidation formula cannot exceed 255 characters");
                }

                _formula = value;

                if (!string.IsNullOrEmpty(value))
                {
                    ResetValue();
                    State = FormulaState.Formula;

                    if (_formula.Any(x => char.IsLetter(x)))
                    {
                        if (RefersToOtherWorksheet(_formula))
                        {
                            var e = new OnFormulaChangedEventArgs();
                            e.isExt = true;
                            _handler.Invoke(e);
                        }
                    }
                }
            }
        }

        private bool RefersToOtherWorksheet(string address)
        {
            if (!string.IsNullOrEmpty(address) && ExcelCellBase.IsValidAddress(address))
            {
                var adr = new ExcelAddress(address);
                return !string.IsNullOrEmpty(adr.WorkSheetName) && adr.WorkSheetName != _workSheetName;
            }
            else if (!string.IsNullOrEmpty(address))
            {
                var tokens = SourceCodeTokenizer.Default.Tokenize(address, _workSheetName);
                if (!tokens.Any()) return false;
                var addressTokens = tokens.Where(x => x.TokenTypeIsSet(TokenType.ExcelAddress));
                foreach (var token in addressTokens)
                {
                    var adr = new ExcelAddress(token.Value);
                    if (!string.IsNullOrEmpty(adr.WorkSheetName) && adr.WorkSheetName != _workSheetName)
                        return true;
                }

            }
            return false;
        }

        internal abstract void ResetValue();

        /// <summary>
        /// This value will be stored in the xml. Can be overridden by subclasses
        /// </summary>
        internal virtual string GetXmlValue()
        {
            if (State == FormulaState.Formula)
            {
                return ExcelFormula;
            }
            return GetValueAsString();
        }

        /// <summary>
        /// Returns the value as a string. Must be implemented by subclasses
        /// </summary>
        /// <returns></returns>
        protected abstract string GetValueAsString();
    }
}
