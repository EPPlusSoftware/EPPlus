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
using OfficeOpenXml.Encryption;
using OfficeOpenXml.Utils;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.RichData.Types
{
    internal class ExcelRichTypeValueType
    {
        public ExcelRichTypeValueType(string name) 
        {
            _name = name;
        }

        public readonly string _name;

        public string Name => _name;

        public string ExtLstXml { get; set; }

        public List<ExcelRichTypeValueKey> _keyFlags = new List<ExcelRichTypeValueKey>();

        public List<ExcelRichTypeValueKey> KeyFlags = new List<ExcelRichTypeValueKey>();

        internal void WriteXml(StreamWriter sw)
        {
            sw.Write($"<type name=\"{Name.EncodeXMLAttribute()}\">");
            sw.Write("<keyFlags>");
            foreach (var flag in KeyFlags)
            {
                flag.WriteXml(sw);
            }
            sw.Write("</keyFlags>");
            if(!string.IsNullOrEmpty(ExtLstXml))
            {
                sw.Write(ExtLstXml);
            }
            sw.Write("</type>");
        }
    }
}
