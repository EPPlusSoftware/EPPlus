using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;

namespace OfficeOpenXml.RichData.Types
{
    internal class ExcelRichTypeValueKey
    {
        public ExcelRichTypeValueKey(string name)
        {
            Name = name;
        }
        public string Name { get; set; }
        public RichValueKeyFlags Flags { get; set; }

        internal void WriteXml(StreamWriter sw)
        {
            sw.Write($"<key name=\"{Name}\">");
            foreach(var flag in GetEnumFlags(Flags))
            {
                sw.Write($"<flag name=\"{flag}\" value=\"1\" />");
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
