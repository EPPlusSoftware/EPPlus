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
        /// 
        /// </summary>
        public int Position { get; set; } = -1;
        /// <summary>
        /// 
        /// </summary>
        public int Length { get; set; } = 0;
        /// <summary>
        /// 
        /// </summary>
        public eDataTypes DataType { get; set; } = eDataTypes.Unknown;
        /// <summary>
        /// 
        /// </summary>
        public char PaddingCharacter { get; set; } = ' ';
        /// <summary>
        /// 
        /// </summary>
        public PaddingAlignmentType PaddingType { get; set; } = PaddingAlignmentType.Auto;
        /// <summary>
        /// Force writing to file, this will only write the n first found characters, where n is column width
        /// </summary>
        public bool ForceWrite { get; set; } = false;
        /// <summary>
        /// 
        /// </summary>
        public bool UseColumn { get; set; } = true;

    }
}
