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
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml
{
    /// <summary>
    /// Describes how to split a fixed width text. Used by the ExcelRange.LoadFromText method
    /// </summary>
    public class ExcelTextFormatFixedWidth : ExcelTextFormatFixedWidthBase
    {
        /// <summary>
        /// 
        /// </summary>
        public ExcelTextFormatFixedWidth() : base()
        {
        }
    }
}
