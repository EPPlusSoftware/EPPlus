/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  10/07/2025         EPPlus Software AB       Initial release EPPlus 8.3
 *************************************************************************************************/
namespace OfficeOpenXml.Data.Connection
{
    /// <summary>
    /// Represents a parameter for an Excel connection.
    /// </summary>
    public class ExcelConnectionParameters
    {
        /// <summary>
        /// The name of the parameter.
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// The SQL type of the parameter. Defaults to 0.
        /// </summary>
        public int SqlType { get; set; } = 0;

        /// <summary>
        /// The type of the parameter. Defaults to Prompt.
        /// </summary>
        public eConnectionParameterType ParameterType { get; set; } = eConnectionParameterType.Prompt;

        /// <summary>
        /// Indicates whether to refresh on change. Defaults to false.
        /// </summary>
        public bool RefreshOnChange { get; set; } = false;

        /// <summary>
        /// Prompt text for the parameter.
        /// </summary>
        public string Prompt { get; set; }

        /// <summary>
        /// Boolean value for the parameter.
        /// </summary>
        public bool? Boolean { get; set; }

        /// <summary>
        /// Double value for the parameter.
        /// </summary>
        public double? Double { get; set; }

        /// <summary>
        /// Integer value for the parameter.
        /// </summary>
        public int? Integer { get; set; }

        /// <summary>
        /// String value for the parameter.
        /// </summary>
        public string String { get; set; }

        /// <summary>
        /// Cell reference for the parameter.
        /// </summary>
        public string Cell { get; set; }
    }

    /// <summary>
    /// Example enum for parameter types. Replace or extend as needed.
    /// </summary>
    public enum eConnectionParameterType
    {
        Prompt,
        Value,
        Cell
    }
}