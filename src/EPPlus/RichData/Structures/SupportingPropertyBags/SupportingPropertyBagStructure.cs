/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  11/11/2024         EPPlus Software AB       Initial release EPPlus 8
 *************************************************************************************************/
using OfficeOpenXml.Utils;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.RichData.Structures.SupportingPropertyBags
{
    [DebuggerDisplay("keys: {Keys.Count}")]
    internal class SupportingPropertyBagStructure
    {
        public SupportingPropertyBagStructure(List<ExcelRichValueStructureKey> keys)
        {
            Keys = keys;
        }

        internal string Type { get; }
        internal List<ExcelRichValueStructureKey> Keys { get; }

        internal void WriteXml(StreamWriter sw)
        {
            sw.Write($"<s>");
            foreach (var key in Keys)
            {
                sw.Write($"<k n=\"{key.Name.EncodeXMLAttribute()}\" {GetTypeAttribute(key)}/>");
            }
            sw.Write("</s>");
        }

        private string GetTypeAttribute(ExcelRichValueStructureKey key)
        {
            if (key.DataType != RichValueDataType.Decimal)
            {
                return $"t =\"{key.GetDataTypeString()}\"";
            }
            return "";
        }
    }
}
