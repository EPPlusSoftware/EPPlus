/*************************************************************************************************
 Required Notice: Copyright (C) EPPlus Software AB. 
 This software is licensed under PolyForm Noncommercial License 1.0.0 
 and may only be used for noncommercial purposes 
 https://polyformproject.org/licenses/noncommercial/1.0.0/

 A commercial license to use this software can be purchased at https://epplussoftware.com
*************************************************************************************************
 Date               Author                       Change
*************************************************************************************************
 03/13/2025         EPPlus Software AB       Initial release EPPlus 8
*************************************************************************************************/
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml
{
    /// <summary>
    /// Represents a #SPILL! error
    /// </summary>
    public class ExcelSpillErrorValue : ExcelRichDataErrorValue
    {
        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="rowOffset"></param>
        /// <param name="colOffset"></param>
        public ExcelSpillErrorValue(int rowOffset, int colOffset) : base(eErrorType.Spill)
        {
            SpillRowOffset = rowOffset;
            SpillColOffset = colOffset;
        }

        internal int SpillRowOffset { get; set; }
        internal int SpillColOffset { get; set; }
        internal bool IsPropagated
        {
            get;
            set;
        }
    }
}
