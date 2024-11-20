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
using OfficeOpenXml.Core.CellStore;
using OfficeOpenXml.RichData.IndexRelations;
using OfficeOpenXml.Utils;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml;

namespace OfficeOpenXml.RichData.RichValueArrays
{
    internal class ExcelRichDataArrayValue : IndexEndpoint
    {
        public ExcelRichDataArrayValue(ExcelRichData richData, RichDataIndexStore store, XmlReader xr) : base(store, RichDataEntities.RichDataArrayValue)
        {
            _richData = richData;
            ReadXml(xr);
        }

        private readonly ExcelRichData _richData;

        public ExcelRichDataArrayValueType ValueType
        {
            get; private set;
        }

        public string Value { get; set; }

        public void ReadXml(XmlReader xr)
        {
            do
            {
                if (xr.IsElementWithName("v"))
                {
                    var t = xr.GetAttribute("t");
                    ValueType = ToValueType(t);
                    xr.Read();
                    if(ValueType == ExcelRichDataArrayValueType.RichValue)
                    {
                        var rvIx = int.Parse(xr.Value);
                        var rvId = _richData.Values.GetIdByIndex(rvIx);
                        Value = rvId.ToString();
                    }
                    else
                    {
                        Value = xr.Value;
                    }
                }
                else if (xr.IsEndElementWithName("v"))
                {
                    break;
                }
            }
            while (xr.Read());
        }

        // 2.7.32 ST_ArrayValueType
        private ExcelRichDataArrayValueType ToValueType(string t)
        {
            switch(t)
            {
                case "d":
                    return ExcelRichDataArrayValueType.RealNumber;
                case "i":
                    return ExcelRichDataArrayValueType.Integer;
                case "b":
                    return ExcelRichDataArrayValueType.Boolean;
                case "e":
                    return ExcelRichDataArrayValueType.Error;
                case "s":
                    return ExcelRichDataArrayValueType.Text;
                case "r":
                    return ExcelRichDataArrayValueType.RichValue;
                case "a":
                    return ExcelRichDataArrayValueType.Array;
                default:
                    throw new ArgumentException($"Invalid rich data array value type: {t}");
            }
        }

        private string GetValueTypeForXml()
        {
            switch(ValueType)
            {
                case ExcelRichDataArrayValueType.RealNumber:
                    return "d";
                case ExcelRichDataArrayValueType.Integer:
                    return "i";
                case ExcelRichDataArrayValueType.Boolean:
                    return "b";
                case ExcelRichDataArrayValueType.Error:
                    return "e";
                case ExcelRichDataArrayValueType.Text:
                    return "s";
                case ExcelRichDataArrayValueType.RichValue:
                    return "r";
                case ExcelRichDataArrayValueType.Array:
                    return "a";
                default:
                    return "s";
            }
        }

        internal void WriteXml(StreamWriter sw)
        {
            var vt = GetValueTypeForXml();
            var val = Value;
            if(ValueType == ExcelRichDataArrayValueType.RichValue)
            {
                var rvId = uint.Parse(Value);
                var rvIx = _richData.Values.GetIndexById(rvId);
                val = rvIx.ToString();
            }
            sw.Write($"<v t=\"{vt}\">{val}</v>");
        }
    }
}
