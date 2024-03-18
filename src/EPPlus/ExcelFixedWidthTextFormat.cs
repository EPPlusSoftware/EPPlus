﻿/*************************************************************************************************
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
using System.Collections.Generic;
using System.Text;
using System.Globalization;

namespace OfficeOpenXml
{

    /// <summary>
    /// 
    /// </summary>
    public class ExcelFixedWidthTextFormat : ExcelTextFormat
    {
        /// <summary>
        /// 
        /// </summary>
        public int[] ColumnLengths;

        /// <summary>
        /// 
        /// </summary>
        public ExcelFixedWidthTextFormat() : base()
        {
            DataTypes = null;
            ColumnLengths = null;
        }
        public ExcelFixedWidthTextFormat(params int[] columnLengths) : base()
        {
            DataTypes = null;
            ColumnLengths = columnLengths;
        }
    }

    /// <summary>
    /// 
    /// </summary>
    public class ExcelFixedWidthTextOutputFormat : ExcelTextFormatBase
    {
        /// <summary>
        /// 
        /// </summary>
        public ExcelFixedWidthTextOutputFormat() : base()
        {

        }
    }
}
