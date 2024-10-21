using OfficeOpenXml.FormulaParsing.Excel.Functions.Information;
using OfficeOpenXml.RichData.IndexRelations;
using OfficeOpenXml.RichData.Mappings;
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
    internal abstract class ExcelRichValue : IndexEndpoint
    {
        public ExcelRichValue(RichDataIndexStore store, ExcelRichData richData, RichDataStructureTypes structureType)
            : base(store, RichDataEntities.RichValue)
        {
            //_workbook = workbook;
            //StructureId = workbook.RichData.Structures.GetStructureId(structureType);
            //Structure = _workbook.RichData.Structures.StructureItems[StructureId];
            var structure = richData.Structures.GetByType(structureType);
            StructureId = structure.Id;
            Structure = structure;
            _richData = richData;
            _indexStore = store;
            As = new ExcelRichValueAsType(this);
            richData.Structures.CreateRelation(this, structure, IndexType.ZeroBasedPointer);
        }


        private readonly ExcelRichData _richData;
        private readonly RichDataIndexStore _indexStore;
        public uint StructureId { get; set; }
        public ExcelRichValueStructure Structure { get; set; }
        //public List<string> Values { get; } = new List<string>();

        public ExcelRichValueAsType As { get; private set; }

        private Dictionary<string, string> _keysAndValues = new Dictionary<string, string>();

        private Dictionary<string, IndexRelation> _relations = new Dictionary<string, IndexRelation>();

        public RichValueFallbackType FallbackType { get; internal set; } = RichValueFallbackType.Decimal;

        public string FallbackValue { get; set; }

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
                sw.Write($"<v>{ConvertUtil.ExcelEscapeString(_keysAndValues[key])}</v>");
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

        public override void DeleteMe()
        {
            base.DeleteMe();
            foreach(var key in Structure.Keys)
            {
                if(key.IsRelation)
                {
                    DeleteRelation(key.Name);
                }
            }
        }

        public Uri GetRelation(string key)
        {
            return GetRelation(key, out IndexRelation relIx);
        }

        public bool DeleteRelation(string key)
        {
            if (!_relations.ContainsKey(key)) return false;
            var rel = _relations[key];
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