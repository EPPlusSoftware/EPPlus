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
using OfficeOpenXml.RichData.Structures.Constants;
using OfficeOpenXml.Utils;
using OfficeOpenXml.Utils.Extensions;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;

namespace OfficeOpenXml.RichData.Structures
{
    [DebuggerDisplay("Type: {Type}")]
    internal abstract class ExcelRichValueStructure : IndexEndpoint
    {
        public ExcelRichValueStructure(string typeName, List<ExcelRichValueStructureKey> keys, RichDataIndexStore store)
            : base(store, RichDataEntities.RichStructure)
        {
            Type = typeName;
            Keys = keys;
        }

        public abstract RichDataStructureTypes StructureType { get; }

        internal List<ExcelRichValueStructureKey> Keys { get; private set; }

        public string Type { get; private set; }

        internal void WriteXml(StreamWriter sw)
        {
            sw.Write($"<s t=\"{Type.EncodeXMLAttribute()}\">");
            foreach (var key in Keys)
            {
                sw.Write($"<k n=\"{key.Name.EncodeXMLAttribute()}\" {GetTypeAttribute(key)}/>");
            }
            sw.Write("</s>");
        }

        private string GetTypeAttribute(ExcelRichValueStructureKey key)
        {
            if (key.DataType != RichValueDataType.Decimal)
            {
                return $"t =\"{key.GetDataTypeString()}\"";
            }
            return "";
        }

        internal bool IsValidKey(string structureKey)
        {
            return Keys.Any(x => x.Name == structureKey);
        }

        /// <summary>
        /// Returns a list of indexes that refers to Keys
        /// that are representing a relation (key starts with _rvRel:
        /// </summary>
        /// <returns></returns>
        internal List<int> GetRelationIndexes()
        {
            var result = new List<int>();
            for(var i = 0; i < Keys.Count; i++)
            {
                if (Keys[i].Name.StartsWith($"{SpecialKeyNames.Prefixes.RvRel}:"))
                {
                    result.Add(i);
                }
            }
            return result;
        }

        internal int? GetFirstRelationIndex()
        {
            var indexes = GetRelationIndexes();
            if(indexes.Count > 0)
            {
                return indexes.First();
            }
            return null;
        }

        /// <summary>
        /// Returns the 0-based index of a key that is a Rich Value relation and its property Name is equal to <paramref name="name"/>.
        /// </summary>
        /// <param name="name"></param>
        /// <returns>index of the found key or null if no such key exists</returns>
        internal int? GetRelationIndexByName(string name)
        {
            for(var i = 0; i < Keys.Count;i++)
            {
                var key = Keys[i];
                if(key.IsRelation && key.RelationName == name)
                {
                    return i;
                }
            }
            return null;
        }

        public override void OnConnectedEntityDeleted(ConnectedEntityDeletedArgs e)
        {
            base.OnConnectedEntityDeleted(e);
            // delete the structure if there is only one rich value left pointing at it
            // if there is only 1 it means it is the last value that is about to be deleted.
            if(e.DeletedEntity.EntityType == RichDataEntities.RichValue)
            {
                var rels = GetIncomingRelations(x => x.From.EntityType == RichDataEntities.RichValue);
                if(rels.Count() <= 1)
                {
                    DeleteMe();
                }
            }
        }
    }
}
