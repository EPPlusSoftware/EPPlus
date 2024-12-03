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
using OfficeOpenXml.RichData.IndexRelations;
using OfficeOpenXml.RichData.Structures;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.RichData.RichValues
{
    internal class ExcelRichValueValue : IndexEndpoint
    {
        public ExcelRichValueValue(ExcelRichValueStructureKey key, string value, RichDataIndexStore store) : base(store, RichDataEntities.RichValueValue)
        {
            _key = key;
            Value = value;
        }

        private readonly ExcelRichValueStructureKey _key;

        public ExcelRichValueStructureKey Key => _key;

        public string Value { get; set; }

        public int? ValueInt
        {
            get
            {
                if(int.TryParse(Value, out int valueInt))
                {
                    return valueInt;
                }
                return null;
            }
        }

        public uint? ValueUint
        {
            get
            {
                if (uint.TryParse(Value, out uint valueUint))
                {
                    return valueUint;
                }
                return null;
            }
        }

        public double? ValueDouble
        {
            get
            {
                try
                {
                    return double.Parse(Value, CultureInfo.InvariantCulture);
                }
                catch
                {
                    return null;
                }
            }
        }

        public bool? ValueBool
        {
            get
            {
                var i = ValueInt;
                if(i.HasValue)
                {
                    return i.Value == 1;
                }
                return null;
            }
        }
    }
}
