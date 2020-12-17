/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  12/10/2020         EPPlus Software AB       EPPlus 5.5
 *************************************************************************************************/
using OfficeOpenXml.Table;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Attributes
{
    /// <summary>
    /// Use this attribute on a class or an interface to insert a column with a formula
    /// </summary>
    [AttributeUsage(AttributeTargets.Class | AttributeTargets.Interface, AllowMultiple = true)]
    public class EpplusFormulaTableColumnAttribute : EpplusTableColumnAttributeBase
    {
        private string _formula = null;
        private string _formulaR1C1 = null;

        /// <summary>
        /// The spreadsheet formula (don't include the leading '='). If you use the {row} placeholder in the formula it will be replaced with the actual row of each cell in the column.
        /// </summary>
        public string Formula
        {
            get
            {
                return _formula;
            }
            set
            {
                if(!string.IsNullOrEmpty(_formulaR1C1) && !string.IsNullOrEmpty(value))
                {
                    throw new InvalidOperationException("EpplusFormulaTableColumn attribute: Formula cannot be set if FormulaR1C1 is not null or empty.");
                }
                _formula = value;
            }
        }

        /// <summary>
        /// The spreadsheet formula (don't include the leading '=') in R1C1 format.
        /// </summary>
        public string FormulaR1C1
        {
            get
            {
                return _formulaR1C1;
            }
            set
            {
                if (!string.IsNullOrEmpty(_formula) && !string.IsNullOrEmpty(value))
                {
                    throw new InvalidOperationException("EpplusFormulaTableColumn attribute: FormulaR1C1 cannot be set if Formula is not null or empty.");
                }
                _formulaR1C1 = value;
            }
        }
    }
}
