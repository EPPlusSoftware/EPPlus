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
using OfficeOpenXml.RichData.IndexRelations.EventArguments;
using OfficeOpenXml.Utils;
using System;
using System.IO;
using System.Linq;
using System.Xml;

namespace OfficeOpenXml.RichData.RichValueArrays
{
    internal class ExcelRichDataArray : IndexEndpoint
    {
        public ExcelRichDataArray(RichDataDatabase richDataDb, XmlReader xr) : base(richDataDb.IndexStore, RichDataEntities.RichDataArray)
        {
            _richDataDb = richDataDb;
            _indexStore = richDataDb.IndexStore;
            _values = new IndexedSubsetCollection<ExcelRichDataArrayValue>(_richDataDb.RichDataArrayValues);
            ReadXml(xr);
        }

        private readonly RichDataIndexStore _indexStore;
        private readonly RichDataDatabase _richDataDb;
        private readonly IndexedSubsetCollection<ExcelRichDataArrayValue> _values;

        public uint RichValueId { get; set; }

        public int Rows { get; set; }

        public int? Columns { get; set; }

        public IndexedSubsetCollection<ExcelRichDataArrayValue> Values => _values;

        private void ReadXml(XmlReader xr)
        {
            var r = xr.GetAttribute("r");
            Rows = int.Parse(r);
            var c = xr.GetAttribute("c");
            if(!string.IsNullOrEmpty(c))
            {
                Columns = int.Parse(c);
            }
            while(xr.Read())
            {
                if(xr.IsElementWithName("v"))
                {
                    var val = new ExcelRichDataArrayValue(_richDataDb, xr);
                    _richDataDb.RichDataArrayValues.Add(val);
                    _values.Add(val);
                }
                else if(xr.IsElementWithName("r"))
                {
                    var rvIxStr = xr.Value;
                    var rvIx = int.Parse(rvIxStr);
                    var rvId = _richDataDb.Values.GetIdByIndex(rvIx);
                    var rv = _richDataDb.Values.Get(rvId);
                    rv.AddRelationTo(this);
                }
                else if(xr.IsEndElementWithName("a"))
                {
                    break;
                }
            }
        }

        internal void WriteXml(StreamWriter sw)
        {
            var rvIx = _richDataDb.Values.GetIndexById(RichValueId);
            if(Columns.HasValue)
            {
                sw.Write($"<a r=\"{Rows}\" c=\"{Columns.Value}\">");
            }
            else
            {
                sw.Write($"<a r=\"{Rows}\">");
            }
            foreach(var val in _values)
            {
                val.WriteXml(sw);
            }
            sw.Write("</a>");
        }

        public override void OnConnectedEntityDeleted(ConnectedEntityDeletedEventArgs e)
        {
            if (Deleted) return;
            base.OnConnectedEntityDeleted(e);
            if (e.DeletedEntity.EntityType == RichDataEntities.RichValue)
            {
                var rels = GetIncomingRelations();
                if (rels.Count() <= 1)
                {
                    // this was the last rich value connected to this relation
                    DeleteMe(e.RelationDeletions);
                }
            }
        }
    }
}
