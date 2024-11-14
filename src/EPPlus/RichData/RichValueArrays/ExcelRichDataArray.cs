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
using OfficeOpenXml.Utils;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml;

namespace OfficeOpenXml.RichData.RichValueArrays
{
    internal class ExcelRichDataArray : IndexEndpoint
    {
        public ExcelRichDataArray(ExcelRichData richData, RichDataIndexStore store, XmlReader xr) : base(store, RichDataEntities.RichDataArray)
        {
            _richData = richData;
            _indexStore = store;
            _values = new IndexedSubsetCollection<ExcelRichDataArrayValue>(richData.RichDataArrayValues);
            ReadXml(xr);
        }

        private readonly RichDataIndexStore _indexStore;
        private readonly ExcelRichData _richData;
        private readonly IndexedSubsetCollection<ExcelRichDataArrayValue> _values;

        public uint RichValueId { get; set; }

        public IndexedSubsetCollection<ExcelRichDataArrayValue> Values => _values;

        private void ReadXml(XmlReader xr)
        {
            var rid = xr.GetAttribute("r");
            RichValueId = uint.Parse(rid);
            while(xr.Read())
            {
                if(xr.IsElementWithName("v"))
                {
                    var val = new ExcelRichDataArrayValue(_richData, _indexStore, xr);
                    _richData.RichDataArrayValues.Add(val);
                    _values.Add(val);
                }
                else if(xr.IsEndElementWithName("a"))
                {
                    break;
                }
            }
        }

        internal void WriteXml(StreamWriter sw)
        {
            var rvIx = _richData.Values.GetIndexById(RichValueId);
            sw.Write($"<a r=\"{rvIx}\">");
            foreach(var val in _values)
            {

            }
            sw.Write("</a>")
        }
    }
}
