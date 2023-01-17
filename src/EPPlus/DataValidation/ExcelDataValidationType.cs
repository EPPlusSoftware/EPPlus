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

    //Va jag vill string -> Class
    //Va måste va string -> enum -> class

    /// <summary>
    /// Types of datavalidation
    /// </summary>
    public class ExcelDataValidationType
    {
        //private ExcelDataValidationType(eDataValidationType validationType, bool allowOperator, string schemaName)
        //{
        //    Type = validationType;
        //    AllowOperator = allowOperator;
        //    SchemaName = schemaName;
        //}

        internal ExcelDataValidationType(eDataValidationType validationType) { Type = validationType; }

        /// <summary>
        /// Validation type
        /// </summary>
        public eDataValidationType Type { get; private set; }
    }
}
