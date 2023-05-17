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
using System.Linq;
using System.Text;
using OfficeOpenXml.DataValidation.Contracts;
using System;
namespace OfficeOpenXml.DataValidation
{
    /// <summary>
    /// Provides functionality for adding datavalidation to a range (<see cref="ExcelRangeBase"/>). Each method will
    /// return a configurable validation.
    /// </summary>
    public interface IRangeDataValidation
    {
        /// <summary>
        /// Adds a <see cref="IExcelDataValidationAny"/> to the range.
        /// </summary>
        /// <returns>A <see cref="ExcelDataValidationAny"/> that can be configured for any validation</returns>
        IExcelDataValidationAny AddAnyDataValidation();
        /// <summary>
        /// Adds a <see cref="ExcelDataValidationInt"/> to the range
        /// </summary>
        /// <returns>A <see cref="ExcelDataValidationInt"/> that can be configured for integer data validation</returns>
        Contracts.IExcelDataValidationInt AddIntegerDataValidation();
        /// <summary>
        /// Adds a <see cref="ExcelDataValidationDecimal"/> to the range
        /// </summary>
        /// <returns>A <see cref="ExcelDataValidationDecimal"/> that can be configured for decimal data validation</returns>
        IExcelDataValidationDecimal AddDecimalDataValidation();
        /// <summary>
        /// Adds a <see cref="ExcelDataValidationDateTime"/> to the range
        /// </summary>
        /// <returns>A <see cref="ExcelDataValidationDecimal"/> that can be configured for datetime data validation</returns>
        IExcelDataValidationDateTime AddDateTimeDataValidation();
        /// <summary>
        /// Adds a <see cref="IExcelDataValidationList"/> to the range
        /// </summary>
        /// <returns>A <see cref="ExcelDataValidationList"/> that can be configured for datetime data validation</returns>
        IExcelDataValidationList AddListDataValidation();
        /// <summary>
        /// Adds a <see cref="ExcelDataValidationInt"/> regarding text length validation to the range.
        /// </summary>
        /// <returns></returns>
        Contracts.IExcelDataValidationInt AddTextLengthDataValidation();
        /// <summary>
        /// Adds a <see cref="IExcelDataValidationTime"/> to the range.
        /// </summary>
        /// <returns>A <see cref="IExcelDataValidationTime"/> that can be configured for time data validation</returns>
        IExcelDataValidationTime AddTimeDataValidation();
        /// <summary>
        /// Adds a <see cref="IExcelDataValidationCustom"/> to the range.
        /// </summary>
        /// <returns>A <see cref="IExcelDataValidationCustom"/> that can be configured for custom validation</returns>
        IExcelDataValidationCustom AddCustomDataValidation();

        /// <summary>
        /// Removes validation from the cell/range
        /// </summary>
        /// <param name="deleteIfEmpty">Delete the validation if it has no more addresses its being applied to. If set to false an <see cref="InvalidOperationException"/> will be thrown if all addresses of a datavalidation has been cleared.</param>
        /// <exception cref="InvalidOperationException">Thrown if <paramref name="deleteIfEmpty"/> is false and all addresses of a datavalidation has been cleared.</exception>
        void ClearDataValidation(bool deleteIfEmpty = false);
    }
}
