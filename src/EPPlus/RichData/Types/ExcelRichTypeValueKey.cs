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
using System.Linq;
using System.Text;
using System.IO;
using OfficeOpenXml.Utils.AttributesUtils;

namespace OfficeOpenXml.RichData.Types
{
    internal class ExcelRichTypeValueKey
    {
        public ExcelRichTypeValueKey(string name)
        {
            Name = name;
            Flags = new List<ExcelRichTypeValueKeyFlag>();
        }
        public string Name { get; set; }
        public List<ExcelRichTypeValueKeyFlag> Flags { get; set; }

        internal void WriteXml(StreamWriter sw)
        {
            sw.Write($"<key name=\"{Name.EncodeXMLAttribute()}\">");
            foreach (var flag in Flags)
            {
                flag.WriteXml(sw);

            }
            sw.Write("</key>");
        }

        private IEnumerable<T> GetEnumFlags<T>(T flags) where T : Enum
        {
            var l=new List<T>();
            var fAll = Convert.ToInt32(flags); 
            foreach (T f in Enum.GetValues(typeof(T)))
            {
                var i = Convert.ToInt32(f);
                if((i & fAll)==i)
                {
                    l.Add(f);
                }
            }
            return l;
        }
    }
}
