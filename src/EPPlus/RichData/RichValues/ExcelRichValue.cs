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
using OfficeOpenXml.RichData.Mappings;
using OfficeOpenXml.RichData.Structures;
using OfficeOpenXml.RichData.Structures.Constants;
using OfficeOpenXml.Utils;
using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;

namespace OfficeOpenXml.RichData.RichValues
{
    internal abstract class ExcelRichValue : IndexEndpoint
    {
        public ExcelRichValue(RichDataDatabase richDataDb, RichDataStructureTypes structureType)
            : base(richDataDb.IndexStore, RichDataEntities.RichValue)
        {
            _richDataDb = richDataDb;
            _indexStore = richDataDb.IndexStore;
            _structureType = structureType;
            Values = new IndexedSubsetCollection<ExcelRichValueValue>(richDataDb.RichValueValues);
            As = new ExcelRichValueAsType(this);
        }

        public ExcelRichValue(RichDataDatabase richDatadb, IndexedSubsetCollection<ExcelRichValueValue> values, RichDataStructureTypes structureType)
            : this(richDatadb, structureType)
        {
            Values = values;
        }


        private readonly RichDataDatabase _richDataDb;
        private readonly RichDataIndexStore _indexStore;
        private readonly RichDataStructureTypes _structureType;
        protected RichDataDatabase RichDataDb => _richDataDb;
        public uint StructureId { get; set; }
        public ExcelRichValueStructure Structure { get; set; }

        public RichDataStructureTypes StructureType => _structureType;

        public ExcelRichValueAsType As { get; private set; }

        public IndexedSubsetCollection<ExcelRichValueValue> Values { get; private set; }

        private Dictionary<string, IndexRelation> _relations = new Dictionary<string, IndexRelation>();

        public RichValueFallbackType FallbackType { get; internal set; } = RichValueFallbackType.Decimal;

        public string FallbackValue { get; set; }

        internal virtual void PostProcessInitialRead()
        {
            for(var ix = 0; ix < Values.Count; ix++)
            {
                var val = Values[ix];
                var intVal = val.ValueInt;
                if(val.Key.DataType == RichValueDataType.RichValue && intVal.HasValue)
                {
                    var refRv = _indexStore.GetItemByIndex(intVal.Value, RichDataEntities.RichValue);
                    val.Value = refRv.Id.ToString();
                    val.AddRelationTo(refRv);
                }
            }
        }

        internal virtual void SetStructure(RichDataDatabase richDataDb)
        {
            var keyNames = Values.Where(v => !string.IsNullOrEmpty(v.Value)).Select(v => v.Key.Name).ToList();
            var existingStructure = Structure;
            var existingStructureRel = default(IndexRelation);
            if(existingStructure != null)
            {
                existingStructureRel = GetOutgoingRelations(x => x.To.EntityType == RichDataEntities.RichStructure).FirstOrDefault();
            }
            Structure = richDataDb.Structures.GetByType(_structureType, keyNames);
            StructureId = Structure.Id;
            if(existingStructure != null && existingStructure.Id != StructureId && existingStructureRel != null)
            {
                existingStructure.OnConnectedEntityDeleted(new ConnectedEntityDeletedEventArgs(this, existingStructureRel, _indexStore, new RelationDeletions(_indexStore)));
                AddRelationTo(Structure);
            }
        }

        internal void WriteXml(StreamWriter sw)
        {
            var id = _richDataDb.Structures.GetIndexById(StructureId);
            // TODO: check this, id should not be null
            if (!id.HasValue) return;
            sw.Write($"<rv s=\"{id}\">");
            if(!string.IsNullOrEmpty(FallbackValue))
            {
                if (FallbackType != RichValueFallbackType.Decimal)
                {
                    sw.Write($"<fb t=\"{GetFallbackAsString()}\">");
                }
                else
                {
                    sw.Write("<fb>");
                }
                sw.Write(FallbackValue);
                sw.Write("</fb>");
            }
            foreach(var key in Structure.Keys)
            {
                var val = Values.FirstOrDefault(x => x.Key.Name == key.Name);
                if(val != null)
                {
                    if (_relations.ContainsKey(val.Key.Name))
                    {
                        var relation = _relations[val.Key.Name];
                        if (relation.To.EntityType == RichDataEntities.RichValueRel)
                        {
                            var relIx = _richDataDb.RichValueRels.GetIndexById(relation.To.Id);
                            sw.Write($"<v>{relIx}</v>");
                        }
                        else
                        {
                            var relIx = _richDataDb.Values.GetIndexById(relation.To.Id);
                            sw.Write($"<v>{relIx}</v>");

                        }
                    }
                    else if (val.Key.DataType == RichValueDataType.RichValue)
                    {
                        var rvId = val.ValueUint;
                        if (rvId.HasValue)
                        {
                            var rvIx = _richDataDb.Values.GetIndexById(rvId.Value);
                            sw.Write($"<v>{rvIx}</v>");

                        }
                    }
                    else if (val.Key.Name == StructureKeyNames.WebImage.WebImageIdentifier)
                    {
                        var imageIx = _richDataDb.WebImages.GetIndexById(val.ValueUint.Value);
                        sw.Write($"<v>{imageIx}</v>");
                    }
                    else if (!string.IsNullOrEmpty(val.Value))
                    {
                        sw.Write($"<v>{ConvertUtil.ExcelEscapeString(val.Value)}</v>");
                    }
                }
            }
            sw.Write("</rv>");
        }
        private string GetFallbackAsString()
        {
            switch (FallbackType)
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

        public void SetRelation(string key, string relationName, Uri relUri, out uint rvRelId)
        {
            int? index;
            List<ExcelRichValueStructureKey> keys;
            if(Structure != null)
            {
                index = Structure.GetRelationIndex(relationName);
                keys = Structure.Keys;
            }
            else
            {
                keys = StructureKeys.GetDefaultKeysByType(_structureType);
                index = ExcelRichValueStructure.GetRelationIndexByName(relationName, keys);
            }
            if (!index.HasValue)
            {
                throw new InvalidOperationException($"Cannot create a relation from structure {Structure.Type}/{Structure.StructureType}");
            }
            var rel = keys[index.Value].Name;
            var relationshipType = RichValueRelationMappings.GetSchema(rel);
            var rvRel = _richDataDb.RichValueRels.AddItem(relUri, relationshipType, this, out IndexRelation r);
            rvRelId = rvRel.Id;
            _relations.Add(key, r);
        }

        /// <summary>
        /// Deletes an entity and its relations
        /// </summary>
        /// <param name="relDeletions">Should be null when calling from classes outside the IndexRelation structure</param>
        public override void DeleteMe(RelationDeletions relDeletions = null)
        {
            base.DeleteMe(relDeletions);
            foreach(var key in Structure.Keys)
            {
                if(key.IsRelation)
                {
                    DeleteRelation(key.Name, relDeletions);
                }
            }
        }

        public Uri GetRelation(string key)
        {
            return GetRelation(key, out IndexRelation relIx);
        }

        private bool DeleteRelation(string key,  RelationDeletions relDeletions)
        {
            if (!_relations.ContainsKey(key)) return false;
            var rel = _relations[key];
            var e = new ConnectedEntityDeletedEventArgs(this, rel, _indexStore, relDeletions);
            rel.To.OnConnectedEntityDeleted(e);
            return _indexStore.DeleteRelation(rel);
        }

        public Uri GetRelation(string key, out IndexRelation indexRelation)
        {
            indexRelation = null;
            if(_relations.ContainsKey(key))
            {
                indexRelation = _relations[key];
                var rdRel = _richDataDb.RichValueRels.GetItem(indexRelation.To.Id);
                return rdRel.TargetUri;
            }
            return null;
        }

        public virtual void SetValue(string key, string value)
        {
            var val = Values.FirstOrDefault(x => x.Key.Name == key);
            if(val == null)
            {
                var k = StructureKeys.GetKey(_structureType, key);
                val = new ExcelRichValueValue(k, value, _indexStore);
                _richDataDb.RichValueValues.Add(val);
                Values.Add(val);
            }
            else
            {
                val.Value = value;
            }
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

        protected void SetValue(string key, double? value)
        {
            SetValue(key, value.HasValue ? value.Value.ToString(CultureInfo.InvariantCulture) : null);
        }


        protected void SetValue(string key, bool? value)
        {
            if (value.HasValue)
            {
                SetValue(key, value.Value ? 1 : 0);
            }
            else
            {
                SetValue(key, string.Empty);
            }
        }

        public string GetValue(string key)
        {
            var val = Values.FirstOrDefault(x => x.Key.Name == key);
            return val?.Value;
        }

        protected int? GetValueInt(string key)
        {
            var val = Values.FirstOrDefault(x => x.Key.Name == key);
            return val?.ValueInt;
        }

        protected double? GetValueDouble(string key)
        {
            var val = Values.FirstOrDefault(x => x.Key.Name == key);
            return val?.ValueDouble;
        }

        protected bool? GetValueBool(string key)
        {
            var val = Values.FirstOrDefault(x => x.Key.Name == key);
            return val?.ValueBool;
        }

        public override void OnConnectedEntityDeleted(ConnectedEntityDeletedEventArgs e)
        {
            base.OnConnectedEntityDeleted(e);
            if(e.DeletedEntity.EntityType == RichDataEntities.FutureMetadataRichDataBlock)
            {
                var rels = GetIncomingRelations(x => x.From.EntityType == RichDataEntities.FutureMetadataRichDataBlock);
                if(rels.Count() <= 1)
                {
                    DeleteMe(e.RelationDeletions);
                }
            }
        }
    }
}