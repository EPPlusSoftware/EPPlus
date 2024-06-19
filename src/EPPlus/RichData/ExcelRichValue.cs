using OfficeOpenXml.Utils;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace OfficeOpenXml.RichData
{
    internal class ExcelRichValue
    {
        public ExcelRichValue(int structureId)
        {
            StructureId = structureId;
        }

        public int StructureId { get; set; }
        public ExcelRichValueStructure Structure { get; set; }
        public List<string> Values { get; }=new List<string>();
        public RichValueFallbackType Fallback { get; internal set; } = RichValueFallbackType.Decimal;

        internal void WriteXml(StreamWriter sw)
        {
            sw.Write($"<rv s=\"{StructureId}\">");
            if (Fallback != RichValueFallbackType.Decimal)
            {
                sw.Write($"<fb t=\"{GetFallbackAsString()}\" />");
            }
            foreach(var v in Values)
            {
                sw.Write($"<v>{ConvertUtil.ExcelEscapeString(v)}</v>");
            }
            sw.Write("</rv>");
        }
        private string GetFallbackAsString()
        {
            switch (Fallback)
            {
                case RichValueFallbackType.Boolean:
                    return "b";
                case RichValueFallbackType.Error:
                    return "e";
                case RichValueFallbackType.String:
                    return "s";
                default:
                    return "n";
            }
        }
        public void AddSpillError(int rowOffset, int colOffset, string subType)
        {            
            foreach(var s in Structure.Keys)
            {
                switch(s.Name)
                {
                    case "colOffset":
                        Values.Add(colOffset.ToString());
                        break;
                    case "rwOffset":
                        Values.Add(rowOffset.ToString());
                        break;
                    case "errorType":
                        Values.Add(RichDataErrorType.Spill);
                        break;
                    case "subType":
                        Values.Add(subType);
                        break;
                }
            }
        }
        public void AddPropagatedError(string errorType, bool propagated)
        {
            foreach (var s in Structure.Keys)
            {
                switch (s.Name)
                {
                    case "errorType":
                        Values.Add(errorType);
                        break;
                    case "propagated":
                        Values.Add(propagated ? "1" : "0");
                        break;
                }
            }
        }
        public void AddError(string errorType, string subType)
        {
            foreach (var s in Structure.Keys)
            {
                switch (s.Name)
                {
                    case "errorType":
                        Values.Add(errorType);
                        break;
                    case "subType":
                        Values.Add(subType);
                        break;
                }
            }
        }
        Dictionary<string, string> _keyValues = null;
        internal bool HasValue(string[] keys, string[] values)
        {
            if(_keyValues==null)
            {
                _keyValues = new Dictionary<string, string>();
                for(int i=0;i < Structure.Keys.Count;i++)
                {
                    _keyValues.Add(Structure.Keys[i].Name, Values[i]);
                }
            }
            
            for(int i=0;i<keys.Length;i++)
            {
                if (_keyValues.TryGetValue(keys[i], out string s)==false || s != values[i])
                {
                    return false;
                }
            }
            return true;
        }
    }
}