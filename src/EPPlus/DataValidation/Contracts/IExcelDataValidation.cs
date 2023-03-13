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

namespace OfficeOpenXml.DataValidation.Contracts
{
    /// <summary>
    /// A generic interface for all data validations. Specialized implementation interfaces should inherit this interface.
    /// </summary>
    public interface IExcelDataValidation
    {
        /// <summary>
        /// Unique id of the data validation
        /// </summary>
        string Uid { get; }
        /// <summary>
        /// Address of data validation
        /// </summary>
        ExcelAddress Address { get; }
        /// <summary>
        /// Validation type
        /// </summary>
        ExcelDataValidationType ValidationType { get; }
        /// <summary>
        /// Controls how Excel will handle invalid values.
        /// </summary>
        ExcelDataValidationWarningStyle ErrorStyle { get; set; }
        /// <summary>
        /// True if input message should be shown
        /// </summary>
        bool? AllowBlank { get; set; }
        /// <summary>
        /// True if input message should be shown
        /// </summary>
        bool? ShowInputMessage { get; set; }
        /// <summary>
        /// True if error message should be shown.
        /// </summary>
        bool? ShowErrorMessage { get; set; }
        /// <summary>
        /// Title of error message box (see property ShowErrorMessage)
        /// </summary>
        string ErrorTitle { get; set; }
        /// <summary>
        /// Error message box text (see property ShowErrorMessage)
        /// </summary>
        string Error { get; set; }
        /// <summary>
        /// Title of info box if input message should be shown (see property ShowInputMessage)
        /// </summary>
        string PromptTitle { get; set; }
        /// <summary>
        /// Info message text (see property ShowErrorMessage)
        /// </summary>
        string Prompt { get; set; }
        /// <summary>
        /// True if the current validation type allows operator.
        /// </summary>
        bool AllowsOperator { get; }
        /// <summary>
        /// Validates the state of the validation.
        /// </summary>
        void Validate();

        /// <summary>
        /// Use this property to cast an instance of <see cref="IExcelDataValidation"/> to its subtype, see <see cref="ExcelDataValidationAsType"/>.
        /// </summary>
        ExcelDataValidationAsType As { get; }

        /// <summary>
        /// Defines mode for Input Method Editor used in east-asian languages
        /// </summary>
        ExcelDataValidationImeMode ImeMode { get; set; }

        /// <summary>
        /// Indicates whether this instance is stale, see https://github.com/EPPlusSoftware/EPPlus/wiki/Data-validation-Exceptions
        /// </summary>
        bool IsStale { get; }



    }
}
