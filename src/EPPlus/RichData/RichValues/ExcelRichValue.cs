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
using OfficeOpenXml.FormulaParsing.Excel.Functions.Information;
using OfficeOpenXml.FormulaParsing.Excel.Functions.MathFunctions;
using OfficeOpenXml.RichData.IndexRelations;
using OfficeOpenXml.RichData.IndexRelations.EventArguments;
using OfficeOpenXml.RichData.Mappings;
using OfficeOpenXml.RichData.RichValues.Errors;
using OfficeOpenXml.RichData.Structures;
using OfficeOpenXml.RichData.Structures.Constants;
using OfficeOpenXml.RichData.Structures.LocalImages;
using OfficeOpenXml.Utils;
using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;

namespace OfficeOpenXml.RichData.RichValues
{
    internal abstract class ExcelRichValue : IndexEndpoint
    {
        public ExcelRichValue(RichDataIndexStore store, ExcelRichData richData, RichDataStructureTypes structureType)
            : base(store, RichDataEntities.RichValue)
        {
            //var structure = richData.Structures.GetByType(structureType);
            //StructureId = structure.Id;
            //Structure = structure;
            _richData = richData;
            _indexStore = store;
            _structureType = structureType;
            As = new ExcelRichValueAsType(this);
            //richData.Structures.CreateRelation(this, structure, IndexType.ZeroBasedPointer);
        }


        private readonly ExcelRichData _richData;
        private readonly RichDataIndexStore _indexStore;
        private readonly RichDataStructureTypes _structureType;
        public uint StructureId { get; set; }
        public ExcelRichValueStructure Structure { get; set; }
        //public List<string> Values { get; } = new List<string>();

        public ExcelRichValueAsType As { get; private set; }

        private Dictionary<string, string> _keysAndValues = new Dictionary<string, string>();

        private Dictionary<string, IndexRelation> _relations = new Dictionary<string, IndexRelation>();

        public RichValueFallbackType FallbackType { get; internal set; } = RichValueFallbackType.Decimal;

        public string FallbackValue { get; set; }


        public void InitRelations(ExcelRichValueCollection values)
        {
            for (var ix = 0; ix < Structure.Keys.Count; ix++)
            {
                var key = Structure.Keys[ix];
                // RvRel - relations
                if (key.IsRelation)
                {
                    var rvRelVal = _keysAndValues[key.Name];
                    var rvRelId = _richData.RichValueRels.GetIdByIndex(int.Parse(rvRelVal));
                    var rvRel = _richData.RichValueRels.Get(rvRelId);
                    SetRelation(key.Name, key.RelationName, rvRel.TargetUri);
                }
                // relation to another richvalue by index
                else if(key.DataType == RichValueDataType.RichValue)
                {
                    var rvIndex = int.Parse(_keysAndValues[key.Name]);
                    var targetRv = values[rvIndex];
                    var relation = AddRelationTo(targetRv, IndexType.ZeroBasedPointer);
                    _relations[key.Name] = relation;
                    _keysAndValues[key.Name] = targetRv.Id.ToString();
                }
                else if(Structure.Type == StructureTypes.WebImage && key.Name == StructureKeyNames.WebImage.WebImageIdentifier) 
                {
                    var imgIx = _keysAndValues[key.Name];
                    var imgId = _richData.WebImages.GetIdByIndex(int.Parse(imgIx));
                    var img = _richData.WebImages.Get(imgId);
                    AddRelationTo(img);
                }
            }
        }

        internal void SetStructure(ExcelRichData richData)
        {
            _keysAndValues = _keysAndValues.Where(kvp => !string.IsNullOrEmpty(kvp.Value)).ToDictionary(kvp => kvp.Key, kvp => kvp.Value);
            var keyNames = _keysAndValues.Select(kvp => kvp.Key).ToList();
            //var hash = ExcelRichValueStructure.CreateKeyHash(keyNames);
            var existingStructure = Structure;
            var existingStructureRel = default(IndexRelation);
            if(existingStructure != null)
            {
                existingStructureRel = GetIncomingRelations(x => x.To.EntityType == RichDataEntities.RichStructure).FirstOrDefault();
            }
            Structure = richData.Structures.GetByType(_structureType, keyNames);
            StructureId = Structure.Id;
            if(existingStructure != null && existingStructure.Id != StructureId && existingStructureRel != null)
            {
                existingStructure.OnConnectedEntityDeleted(new ConnectedEntityDeletedEventArgs(this, existingStructureRel, _indexStore, new RelationDeletions(_indexStore)));
                AddRelationTo(Structure);
            }
        }

        internal void WriteXml(StreamWriter sw)
        {
            var id = _richData.Structures.GetIndexById(StructureId);
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
            foreach (var key in Structure.Keys.ToNameArray())
            {
                if(_relations.ContainsKey(key))
                {
                    var relation = _relations[key];
                    if(relation.To.EntityType == RichDataEntities.RichValueRel)
                    {
                        var relIx = _richData.RichValueRels.GetIndexById(relation.To.Id);
                        sw.Write($"<v>{relIx}</v>");
                    }
                    else
                    {
                        var relIx = _richData.Values.GetIndexById(relation.To.Id);
                        sw.Write($"<v>{relIx}</v>");

                    }
                }
                else
                {
                    sw.Write($"<v>{ConvertUtil.ExcelEscapeString(_keysAndValues[key])}</v>");
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
        #region old code
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
        #endregion

        public void SetRelation(string key, string relationName, Uri relUri)
        {
            var index = Structure.GetRelationIndexByName(relationName);
            if (!index.HasValue)
            {
                throw new InvalidOperationException($"Cannot create a relation from structure {Structure.Type}/{Structure.StructureType}");
            }
            var rel = Structure.Keys[index.Value].Name;
            var relationshipType = RichValueRelationMappings.GetSchema(rel);
            _richData.RichValueRels.AddItem(relUri, relationshipType, this, out IndexRelation r);
            //SetValue(key, relIx);
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
            //relIx = GetValueInt(key);
            //if (!relIx.HasValue) return null;
            //var rdRel = _richData.RichValueRels.Items[relIx.Value];
            //return rdRel.TargetUri;
            indexRelation = null;
            if(_relations.ContainsKey(key))
            {
                indexRelation = _relations[key];
                var rdRel = _richData.RichValueRels.GetItem(indexRelation.To.Id);
                return rdRel.TargetUri;
            }
            return null;
        }

        public void SetValue(string key, string value)
        {
            if (_keysAndValues.ContainsKey(key))
            {
                _keysAndValues.Remove(key);
            }
            if(!string.IsNullOrEmpty(value))
            {
                _keysAndValues[key] = value;
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

        protected double? GetValueDouble(string key)
        {
            if (_keysAndValues.ContainsKey(key))
            {
                try
                {
                    return double.Parse(_keysAndValues[key], CultureInfo.InvariantCulture);
                }
                catch
                {
                    return null;
                }
            }
            return null;
        }

        protected bool? GetValueBool(string key)
        {
            if (_keysAndValues.ContainsKey(key))
            {
                if (int.TryParse(_keysAndValues[key], out var value))
                {
                    return value > 0;
                }
            }
            return null;
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

        #region Old code
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
        #endregion
    }
}