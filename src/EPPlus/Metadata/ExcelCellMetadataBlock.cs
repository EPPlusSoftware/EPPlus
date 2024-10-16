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
using OfficeOpenXml.RichData.IndexRelations;
using OfficeOpenXml.Utils;
using System.Collections.Generic;
using System.Xml;

namespace OfficeOpenXml.Metadata
{
    /// <summary>
    /// Corresponds to a bk-element in the valueMetadata section of the metadata.xml file.
    /// </summary>
    internal class ExcelCellMetadataBlock
    {
        public ExcelCellMetadataBlock()
        {

        }
        public ExcelCellMetadataBlock(XmlReader xr)
        {
            while (xr.IsEndElementWithName("bk") == false && xr.EOF == false)
            {
                if (xr.IsElementWithName("rc"))
                {
                    var t = int.Parse(xr.GetAttribute("t"));
                    var v = int.Parse(xr.GetAttribute("v"));
                    Records.Add(new ExcelCellMetadataRecord(t, v));
                }
                xr.Read();
            }
        }

        public List<ExcelCellMetadataRecord> Records { get; } = new List<ExcelCellMetadataRecord>();
    }
}