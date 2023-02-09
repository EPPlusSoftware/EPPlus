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
using System.Xml;

namespace OfficeOpenXml.DataValidation
{
    /// <summary>
    /// Any value validation.
    /// </summary>
    public class ExcelDataValidationAny : ExcelDataValidation, IExcelDataValidationAny
    {

        internal ExcelDataValidationAny(string uid, string address) : base(uid, address)
        {
        }

        /// <summary>
        /// Constructor for reading data
        /// </summary>
        /// <param name="xr">The XmlReader to read from</param>
        internal ExcelDataValidationAny(XmlReader xr) : base(xr)
        {
        }

        /// <summary>
        /// True if the current validation type allows operator.
        /// </summary>
        public override bool AllowsOperator { get { return false; } }

        /// <summary>
        /// Validation type
        /// </summary>
        public override ExcelDataValidationType ValidationType => new ExcelDataValidationType(eDataValidationType.Any);
    }
}
