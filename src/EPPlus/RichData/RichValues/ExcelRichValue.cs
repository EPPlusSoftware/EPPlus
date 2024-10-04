using OfficeOpenXml.FormulaParsing.Excel.Functions.Information;
using OfficeOpenXml.RichData.RichValues.Errors;
using OfficeOpenXml.RichData.Structures;
using OfficeOpenXml.RichData.Structures.Constants;
using OfficeOpenXml.RichData.Structures.LocalImages;
using OfficeOpenXml.Utils;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;

namespace OfficeOpenXml.RichData.RichValues
{
    internal abstract class ExcelRichValue
    {
        public ExcelRichValue(ExcelWorkbook workbook, RichDataStructureTypes structureType)
        {
            _workbook = workbook;
            StructureId = workbook.RichData.Structures.GetStructureId(structureType);
            Structure = _workbook.RichData.Structures.StructureItems[StructureId];
            As = new ExcelRichValueAsType(this);
        }


        private readonly ExcelWorkbook _workbook;
        public int StructureId { get; set; }
        public ExcelRichValueStructure Structure { get; set; }
        //public List<string> Values { get; } = new List<string>();

        public ExcelRichValueAsType As { get; private set; }

        private Dictionary<string, string> _keysAndValues = new Dictionary<string, string>();

        public RichValueFallbackType Fallback { get; internal set; } = RichValueFallbackType.Decimal;

        internal void WriteXml(StreamWriter sw)
        {
            sw.Write($"<rv s=\"{StructureId}\">");
            if (Fallback != RichValueFallbackType.Decimal)
            {
                sw.Write($"<fb t=\"{GetFallbackAsString()}\" />");
            }
            foreach (var key in Structure.Keys.ToNameArray())
            {
                sw.Write($"<v>{ConvertUtil.ExcelEscapeString(_keysAndValues[key])}</v>");
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

        //private void AddRichValue(bool clearValues, Action action)
        //{
        //    if (clearValues)
        //    {
        //        Values.Clear();
        //    }
        //    action.Invoke();
        //}

        //public void AddSpillError(int rowOffset, int colOffset, string subType, bool clearValues = false)
        //{
        //    AddRichValue(clearValues, () =>
        //    {
        //        foreach (var s in Structure.Keys)
        //        {
        //            switch (s.Name)
        //            {
        //                case "colOffset":
        //                    Values.Add(colOffset.ToString());
        //                    break;
        //                case "rwOffset":
        //                    Values.Add(rowOffset.ToString());
        //                    break;
        //                case "errorType":
        //                    Values.Add(RichDataErrorType.Spill);
        //                    break;
        //                case "subType":
        //                    Values.Add(subType);
        //                    break;
        //            }
        //        }
        //    });

        //}
        //public void AddPropagatedError(string errorType, bool propagated, bool clearValues = false)
        //{
        //    AddRichValue(clearValues, () =>
        //    {
        //        foreach (var s in Structure.Keys)
        //        {
        //            switch (s.Name)
        //            {
        //                case "errorType":
        //                    Values.Add(errorType);
        //                    break;
        //                case "propagated":
        //                    Values.Add(propagated ? "1" : "0");
        //                    break;
        //            }
        //        }
        //    });
        //}

        //public void AddError(string errorType, string subType, bool clearValues = false)
        //{
        //    AddRichValue(clearValues, () =>
        //    {
        //        foreach (var s in Structure.Keys)
        //        {
        //            switch (s.Name)
        //            {
        //                case "errorType":
        //                    Values.Add(errorType);
        //                    break;
        //                case "subType":
        //                    Values.Add(subType);
        //                    break;
        //            }
        //        }
        //    });
        //}

        //public void AddLocalImage(int imageIdentifier, int calcOrigin, string text, bool clearValues = false)
        //{
        //    AddRichValue(clearValues, () =>
        //    {
        //        foreach (var s in Structure.Keys)
        //        {
        //            switch (s.Name)
        //            {
        //                case StructureKeyNames.LocalImages.ImageAltText.RelLocalImageIdentifier:
        //                    Values.Add(imageIdentifier.ToString());
        //                    break;
        //                case StructureKeyNames.LocalImages.ImageAltText.CalcOrigin:
        //                    Values.Add(calcOrigin.ToString());
        //                    break;
        //                case StructureKeyNames.LocalImages.ImageAltText.Text:
        //                    Values.Add(text);
        //                    break;
        //            }
        //        }
        //    });
        //}

        public void SetValue(string key, string value)
        {
            if (Structure.StructureType != RichDataStructureTypes.Preserve && !Structure.IsValidKey(key))
            {
                throw new InvalidOperationException($"Invalid key for rich data of type {Structure.StructureType}: " + key);
            }
            if (_keysAndValues.ContainsKey(key))
            {
                _keysAndValues.Remove(key);
            }
            _keysAndValues[key] = value;
        }

        protected void SetValue(string key, int value)
        {
            SetValue(key, value.ToString());
        }

        protected void SetValue(string key, int? value)
        {
            if(value.HasValue)
            {
                SetValue(key, value.ToString());
            }
            else
            {
                SetValue(key, string.Empty);
            }
        }

        public string GetValue(string key)
        {
            if(_keysAndValues.ContainsKey(key))
            {
                return _keysAndValues[key];
            }
            return string.Empty;
        }

        protected int? GetValueInt(string key)
        {
            if (_keysAndValues.ContainsKey(key))
            {
                if (int.TryParse(_keysAndValues[key], out var value))
                {
                    return value;
                }
            }
            return null;
        }

        //Dictionary<string, string> _keyValues = null;
        //internal bool HasValue(string[] keys, string[] values)
        //{
        //    if (_keyValues == null)
        //    {
        //        _keyValues = new Dictionary<string, string>();
        //        for (int i = 0; i < Structure.Keys.Count; i++)
        //        {
        //            _keyValues.Add(Structure.Keys[i].Name, Values[i]);
        //        }
        //    }

        //    for (int i = 0; i < keys.Length; i++)
        //    {
        //        if (_keyValues.TryGetValue(keys[i], out string s) == false || s != values[i])
        //        {
        //            return false;
        //        }
        //    }
        //    return true;
        //}
    }
}