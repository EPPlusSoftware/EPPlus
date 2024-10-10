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
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.RichData.Structures.SupportingPropertyBags
{
    [DebuggerDisplay("XmlSpace: {XmlSpace}, Value: {Value}")]
    internal class SupportingPropertyBagDataItem
    {
        public SupportingPropertyBagDataItem(string val, string xmlSpace)
        {
            Value = val;
            XmlSpace = xmlSpace;
        }

        public string Value { get; set; }

        public string XmlSpace { get; set; }

        internal void WriteXml(StreamWriter sw)
        {
            if(!string.IsNullOrEmpty(XmlSpace))
            {
                sw.Write($"<v xml:space=\"{XmlSpace}\">{Value}</v>");
            }
            else
            {
                sw.Write($"<v>{Value}</v>");
            }
        }
    }
}
