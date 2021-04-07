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
using System.Xml;
using OfficeOpenXml.Utils;
using OfficeOpenXml.DataValidation.Formulas.Contracts;
using OfficeOpenXml.DataValidation.Events;
using OfficeOpenXml.DataValidation.Exceptions;

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
    internal abstract class ExcelDataValidationFormula : XmlHelper
    {
        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="namespaceManager">Namespacemanger of the worksheet</param>
        /// <param name="topNode">validation top node</param>
        /// <param name="formulaPath">xml path of the current formula</param>
        /// <param name="validationUid">id of the data validation containing this formula</param>
        public ExcelDataValidationFormula(XmlNamespaceManager namespaceManager, XmlNode topNode, string formulaPath, string validationUid)
            : base(namespaceManager, topNode)
        {
            Require.Argument(formulaPath).IsNotNullOrEmpty("formulaPath");
            Require.Argument(validationUid).IsNotNullOrEmpty("validationUid");
            FormulaPath = formulaPath;
            _validationUid = validationUid;
        }

        private string _validationUid;
        private string _formula;
        private List<IFormulaListener> _formulaListeners = new List<IFormulaListener>();


        protected string FormulaPath
        {
            get;
            private set;
        }

        internal void RegisterFormulaListener(IFormulaListener listener)
        {
            _formulaListeners.Add(listener);
        }

        internal void DetachFormulaListener(IFormulaListener listener)
        {
            _formulaListeners.Remove(listener);
        }

        private void OnFormulaChanged(string uid, string oldValue, string newValue) 
        { 
            foreach(var listener in _formulaListeners)
            {
                listener.Notify(new ValidationFormulaChangedArgs { ValidationUid = uid, OldValue = oldValue, NewValue = newValue });
            }
        }

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
                if (!string.IsNullOrEmpty(value))
                {
                    ResetValue();
                    State = FormulaState.Formula;
                }
                if (value != null && MeasureFormulaLength(value) > 255)
                {
                    throw new DataValidationFormulaTooLongException("The length of a DataValidation formula cannot exceed 255 characters");
                }
                var oldValue = _formula;
                _formula = value;
                SetXmlNodeString(FormulaPath, value);
                OnFormulaChanged(_validationUid, oldValue, value);
            }
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
