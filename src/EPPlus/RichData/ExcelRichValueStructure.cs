using System;
using System.Collections.Generic;
using System.IO;

namespace OfficeOpenXml.RichData
{
    internal class ExcelRichValueStructure
    {
        public Dictionary<string, List<List<ExcelRichValueStructureKey>>> RichValueTypes = new Dictionary<string, List<List<ExcelRichValueStructureKey>>>()
        {
            {   "_error",
                new List<List<ExcelRichValueStructureKey>>(){
                    new List<ExcelRichValueStructureKey>()
                    {
                        new ExcelRichValueStructureKey("colOffset", "i"),
                        new ExcelRichValueStructureKey("errorType", "i"),
                        new ExcelRichValueStructureKey("rwOffset", "i"),
                        new ExcelRichValueStructureKey("subType", "i")
                    },
                    new List<ExcelRichValueStructureKey>()
                    {
                        new ExcelRichValueStructureKey("errorType", "i"),
                        new ExcelRichValueStructureKey("propagated", "b")
                    },
                    new List<ExcelRichValueStructureKey>()
                    {
                        new ExcelRichValueStructureKey("errorType", "i"),
                        new ExcelRichValueStructureKey("subType", "b")
                    },
                    new List<ExcelRichValueStructureKey>()
                    {
                        new ExcelRichValueStructureKey("errorType", "i"),
                        new ExcelRichValueStructureKey("targetValue", "r")
                    },
                    new List<ExcelRichValueStructureKey>()
                    {
                        new ExcelRichValueStructureKey("errorType", "i"),
                        new ExcelRichValueStructureKey("field", "s")
                    },
                }
            }
        };
        public string Type { get; set; }
        public List<ExcelRichValueStructureKey> Keys { get;  }=new List<ExcelRichValueStructureKey>();

        internal void WriteXml(StreamWriter sw)
        {
            sw.Write($"<s t=\"{Type}\">");
            foreach(var key in Keys)
            {
                sw.Write($"<k n=\"{key.Name}\" t=\"{key.GetDataTypeString()}\"/>");
            }
            sw.Write("</s>");
        }

        //See MS-XLSX (Extension) 2.3.6.1.3 Error Types for details

        public void SetAsSpillError()
        {
            Type = "_error";
            Keys.AddRange(RichValueTypes[Type][0]);
        }
        public void SetAsPropagatedError()
        {
            Type = "_error";
            Keys.AddRange(RichValueTypes[Type][1]);
        }
        public void SetAsError()
        {
            Type = "_error";
            Keys.AddRange(RichValueTypes[Type][2]);
        }
        public void SetAsBufferError()
        {
            Type = "_error";
            Keys.AddRange(RichValueTypes[Type][3]);
        }
        public void SetAsFieldError()
        {
            Type = "_error";
            Keys.AddRange(RichValueTypes[Type][4]);
        }
    }
}
