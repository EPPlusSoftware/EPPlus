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

namespace OfficeOpenXml.DataValidation
{
    /// <summary>
    /// Enum for available data validation types
    /// </summary>
    public enum eDataValidationType
    {
        /// <summary>
        /// Any value
        /// </summary>
        Any,
        /// <summary>
        /// Integer value
        /// </summary>
        Whole,
        /// <summary>
        /// Decimal values
        /// </summary>
        Decimal,
        /// <summary>
        /// List of values
        /// </summary>
        List,
        /// <summary>
        /// Text length validation
        /// </summary>
        TextLength,
        /// <summary>
        /// DateTime validation
        /// </summary>
        DateTime,
        /// <summary>
        /// Time validation
        /// </summary>
        Time,
        /// <summary>
        /// Custom validation
        /// </summary>
        Custom
    }

    internal static class DataValidationSchemaNames
    {
        public const string Any = "";
        public const string None = "none";
        public const string Whole = "whole";
        public const string Decimal = "decimal";
        public const string List = "list";
        public const string TextLength = "textLength";
        public const string Date = "date";
        public const string Time = "time";
        public const string Custom = "custom";
    }

    /// <summary>
    /// Types of datavalidation
    /// </summary>
    public class ExcelDataValidationType
    {
        internal ExcelDataValidationType(eDataValidationType validationType) { Type = validationType; }

        /// <summary>
        /// Validation type
        /// </summary>
        public eDataValidationType Type { get; private set; }

        /// <summary>
        /// Returns a validation type by <see cref="eDataValidationType"/>
        /// </summary>
        /// <returns>The string output</returns>
        internal string TypeToXmlString()
        {
            switch (Type)
            {
                case eDataValidationType.Any:
                    return DataValidationSchemaNames.Any;
                case eDataValidationType.Whole:
                    return DataValidationSchemaNames.Whole;
                case eDataValidationType.List:
                    return DataValidationSchemaNames.List;
                case eDataValidationType.Decimal:
                    return DataValidationSchemaNames.Decimal;
                case eDataValidationType.TextLength:
                    return DataValidationSchemaNames.TextLength;
                case eDataValidationType.DateTime:
                    return DataValidationSchemaNames.Date;
                case eDataValidationType.Time:
                    return DataValidationSchemaNames.Time;
                case eDataValidationType.Custom:
                    return DataValidationSchemaNames.Custom;
                default:
                    throw new InvalidOperationException("Non supported Validationtype : " + Type.ToString());
            }
        }
    }
}
