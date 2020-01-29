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
using System.Text;
using System.Xml;

namespace OfficeOpenXml
{
    /// <summary>
    /// Error ignore options for a worksheet
    /// </summary>
    public class ExcelIgnoredError : XmlHelper
    {
        internal ExcelIgnoredError(XmlNamespaceManager nsm, XmlNode topNode, ExcelAddressBase address) : base(nsm, topNode)
        {
            SetXmlNodeString("@sqref", address.AddressSpaceSeparated);
        }
        /// <summary>
        /// Ignore errors when numbers are formatted as text or are preceded by an apostrophe
        /// </summary>
        public bool NumberStoredAsText
        {
            get
            {
                return GetXmlNodeBool("@numberStoredAsText");
            }
            set
            {
                SetXmlNodeBool("@numberStoredAsText", value);
            }
        }
        /// <summary>
        /// Calculated Column
        /// </summary>
        public bool CalculatedColumm
        {
            get
            {
                return GetXmlNodeBool("@calculatedColumn");
            }
            set
            {
                SetXmlNodeBool("@calculatedColumn", value);
            }
        }


        /// <summary>
        /// Ignore errors when a formula refers an empty cell
        /// </summary>
        public bool EmptyCellReference
        {
            get
            {
                return GetXmlNodeBool("@emptyCellReference");
            }
            set
            {
                SetXmlNodeBool("@emptyCellReference", value);
            }
        }

        /// <summary>
        /// Ignore errors when formulas fail to Evaluate
        /// </summary>
        public bool EvaluationError
        {
            get
            {
                return GetXmlNodeBool("@evalError");
            }
            set
            {
                SetXmlNodeBool("@evalError", value);
            }
        }
        /// <summary>
        /// Ignore errors when a formula in a region of your worksheet differs from other formulas in the same region.
        /// </summary>
        public bool Formula
        {
            get
            {
                return GetXmlNodeBool("@formula");
            }
            set
            {
                SetXmlNodeBool("@formula", value);
            }
        }
        /// <summary>
        /// Ignore errors when formulas omit certain cells in a region.
        /// </summary>
        public bool FormulaRange
        {
            get
            {
                return GetXmlNodeBool("@formulaRange");
            }
            set
            {
                SetXmlNodeBool("@formulaRange", value);
            }
        }
        /// <summary>
        /// Ignore errors when a cell's value in a Table does not comply with the Data Validation rules specified
        /// </summary>
        public bool ListDataValidation
        {
            get
            {
                return GetXmlNodeBool("@listDataValidation");
            }
            set
            {
                SetXmlNodeBool("@listDataValidation", value);
            }
        }
        /// <summary>
        /// The address
        /// </summary>
        public ExcelAddressBase Address
        {
            get
            {
                return new ExcelAddressBase(GetXmlNodeString("@sqref"));
            }
        }
        /// <summary>
        /// Ignore errors when formulas contain text formatted cells with years represented as 2 digits.
        /// </summary>
        public bool TwoDigitTextYear
        {
            get
            {
                return GetXmlNodeBool("@twoDigitTextYear");
            }
            set
            {
                SetXmlNodeBool("@twoDigitTextYear", value);
            }
        }
        /// <summary>
        /// Ignore errors when unlocked cells contain formulas
        /// </summary>
        public bool UnlockedFormula
        {
            get
            {
                return GetXmlNodeBool("@unlockedFormula");
            }
            set
            {
                SetXmlNodeBool("@unlockedFormula", value);
            }
        }
    }
}
