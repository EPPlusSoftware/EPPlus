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
using OfficeOpenXml.Utils;
using System.Xml;

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
        public ExcelDataValidationFormulaValue(XmlNamespaceManager namespaceManager, XmlNode topNode, string formulaPath)
            : base(namespaceManager, topNode, formulaPath)
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
                SetXmlNodeString(FormulaPath, GetValueAsString());
            }
        }

        internal override void ResetValue()
        {
            Value = default(T);
        }

    }
}
