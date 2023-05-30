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

namespace OfficeOpenXml.DataValidation.Formulas.Contracts
{
    /// <summary>
    /// Interface for a data validation formula
    /// </summary>
    public interface IExcelDataValidationFormula
    {
        /// <summary>
        /// An excel formula 
        /// <para />
        /// Keep in mind that special signs like " and ( Must be made double to work.
        /// <para />
        /// ExcelFormula = "\"Epplus\"" will work. And show up as Epplus in excel
        /// <para />
        /// ExcelFormula = "\"Epplus" will generate a corrupt workbook. As there is no double.
        /// <para />
        /// And "\"\"\"Epplus\"" would show up as "Epplus in excel.
        /// <para />
        /// </summary>
        string ExcelFormula { get; set; }
    }
}
