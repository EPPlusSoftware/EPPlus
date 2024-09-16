/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  07/25/2024         EPPlus Software AB       EPPlus 7
 *************************************************************************************************/
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Metadata.FutureMetadata
{
    internal class ExcelFutureMetadata
    {
        public int Index { get; set; }
        public string Name { get; set; }
        public List<ExcelFutureMetadataType> Types { get; } = new List<ExcelFutureMetadataType>();
        //string _extLstXml;
    }
}
