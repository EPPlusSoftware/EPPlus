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
using System.Linq;
using System.Text;
using System.Xml;

namespace OfficeOpenXml.RichData.RichValueArrays
{
    internal class ExcelRichDataArrayValue : IndexEndpoint
    {
        public ExcelRichDataArrayValue(RichDataIndexStore store, XmlReader xr) : base(store, RichDataEntities.RichDataArrayValue)
        {
            ReadXml(xr);
        }

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
                    Value = xr.Value;
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
    }
}
