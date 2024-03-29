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
using OfficeOpenXml.Utils;
using System.Collections.Generic;
using System.Xml;

namespace OfficeOpenXml.Metadata
{
    internal class ExcelMetadataItem
    {
        public ExcelMetadataItem()
        {

        }
        public ExcelMetadataItem(XmlReader xr)
        {
            while(xr.IsEndElementWithName("bk")==false && xr.EOF==false)
            {
                if(xr.IsElementWithName("rc"))
                {
                    Records.Add(new ExcelMetadataRecord(int.Parse(xr.GetAttribute("t")), int.Parse(xr.GetAttribute("v"))));
                }
                xr.Read();
            }
        }

        public List<ExcelMetadataRecord> Records { get;}= new List<ExcelMetadataRecord>();
    }
}