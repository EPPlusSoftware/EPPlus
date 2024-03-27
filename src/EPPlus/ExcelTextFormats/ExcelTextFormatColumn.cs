/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  12/30/2023         EPPlus Software AB       Initial release EPPlus 7
 *************************************************************************************************/
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;

namespace OfficeOpenXml
{
    /// <summary>
    /// 
    /// </summary>
    public class ExcelTextFormatColumn
    {
        /// <summary>
        /// The start position of the column, is equal to -1 if not set
        /// </summary>
        public int Position { get; set; } = -1;
        /// <summary>
        /// The length of the column.
        /// </summary>
        public int Length { get; set; } = 0;
        /// <summary>
        /// The data type of the column. Is set to Unknown by default
        /// </summary>
        public eDataTypes DataType { get; set; } = eDataTypes.Unknown;
        /// <summary>
        /// The padding type of the column. Is set to auto by default, which will try to pad numbers to the right and strings to the left.
        /// </summary>
        public PaddingAlignmentType PaddingType { get; set; } = PaddingAlignmentType.Auto;
        /// <summary>
        /// Flag to set if column should be used when reading and writing fixed width text.
        /// </summary>
        public bool UseColumn { get; set; } = true;

    }
}
