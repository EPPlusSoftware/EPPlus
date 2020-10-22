/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  10/15/2020         EPPlus Software AB       ToDataTable function
 *************************************************************************************************/
using System;
using System.Collections.Generic;
using System.Text;

namespace OfficeOpenXml.Export.ToDataTable
{
    public class DataColumnMapping
    {
        public int ZeroBasedColumnIndexInRange { get; set; }

        public string DataColumnName { get; set; }

        public Type DataColumnType { get; set; }

        public bool AllowNull { get; set; }

        internal void Validate()
        {
            if(string.IsNullOrEmpty(DataColumnName)) throw new ArgumentNullException("DataColumnName");
            if (ZeroBasedColumnIndexInRange < 0) throw new ArgumentOutOfRangeException("ZeroBasedColumnIndex");
        }
    }
}
