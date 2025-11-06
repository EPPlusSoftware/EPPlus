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
using System;

namespace OfficeOpenXml.Data.Connection
{
    /// <summary>
    /// Represents a text field in a text connection.
    /// </summary>
    public class ExcelConnectionTextField
    {
        /// <summary>
        /// Creates a new text field that can be added to <see cref="ExcelTextProperties.Fields"/>
        /// </summary>
        /// <param name="type">The field data type.</param>
        /// <param name="position">The position. In a fixed index file this is the start position in record.</param>
        public ExcelConnectionTextField(eConnectionTextFieldType type, int position)
        {
            Type = type;
            Position = position;
        }

        /// <summary>
        /// Creates a new text field that can be added to <see cref="ExcelTextProperties.Fields"/>
        /// </summary>
        /// <param name="type">The datatype when importing the field</param>
        public ExcelConnectionTextField(eConnectionTextFieldType type)
        {
            Type= type;
            Position = 0;
        }
        
        public ExcelConnectionTextField(int position)
        {
            Position = position;
        }
        /// <summary>
        /// The format to handle the data type of the field.
        /// </summary>
        public eConnectionTextFieldType Type { get; set; } = eConnectionTextFieldType.General;
        internal int Position { get; set; } = 0;
    }
}