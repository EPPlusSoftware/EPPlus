﻿/*************************************************************************************************
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
namespace OfficeOpenXml.Metadata
{
    internal class ExcelMetadataRecord
    {
        public ExcelMetadataRecord(int recordTypeIndex, int valueTypeIndex)
        {
            RecordTypeIndex= recordTypeIndex;
            ValueTypeIndex = valueTypeIndex;
        }
        public int RecordTypeIndex { get; private set; }
        public int ValueTypeIndex { get; private set; }
    }
}