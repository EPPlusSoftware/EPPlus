/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  01/27/2020         EPPlus Software AB       Initial release EPPlus 5
 *************************************************************************************************/
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Ionic.Zip;
using System.IO;
using System.Security;
using OfficeOpenXml.Packaging.Ionic.Zip;

namespace OfficeOpenXml.Packaging
{
    /// <summary>
    /// A collection of package relationships
    /// </summary>
    public class ZipPackageRelationshipCollection : IEnumerable<ZipPackageRelationship>
    {
        /// <summary>
        /// Relationships dictionary
        /// </summary>
        internal protected Dictionary<string, ZipPackageRelationship> _rels = new Dictionary<string, ZipPackageRelationship>(StringComparer.OrdinalIgnoreCase);
        internal void Add(ZipPackageRelationship item)
        {
            _rels.Add(item.Id, item);
        }
        /// <summary>
        /// Gets the enumerator for the collection
        /// </summary>
        /// <returns>the enumerator</returns>
        public IEnumerator<ZipPackageRelationship> GetEnumerator()
        {
            return _rels.Values.GetEnumerator();
        }

        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return _rels.Values.GetEnumerator();
        }

        internal void Remove(string id)
        {
            _rels.Remove(id);
        }
        internal bool ContainsKey(string id)
        {
            return _rels.ContainsKey(id);
        }
        internal ZipPackageRelationship this[string id]
        {
            get
            {
                return _rels[id];
            }
        }
        internal ZipPackageRelationshipCollection GetRelationshipsByType(string relationshipType)
        {
            var ret = new ZipPackageRelationshipCollection();
            foreach (var rel in _rels.Values)
            {
                if (rel.RelationshipType == relationshipType)
                {
                    ret.Add(rel);
                }
            }
            return ret;
        }

        internal void WriteZip(ZipOutputStream os, string fileName)
        {
            StringBuilder xml = new StringBuilder("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">");
            foreach (var rel in _rels.Values)
            {
                if(rel.TargetUri==null || rel.TargetUri.OriginalString.StartsWith("Invalid:URI"))
                {
                    xml.AppendFormat("<Relationship Id=\"{0}\" Type=\"{1}\" Target=\"{2}\"{3}/>", SecurityElement.Escape(rel.Id), rel.RelationshipType, SecurityElement.Escape(rel.Target), rel.TargetMode == TargetMode.External ? " TargetMode=\"External\"" : "");
                }
                else
                {
                    xml.AppendFormat("<Relationship Id=\"{0}\" Type=\"{1}\" Target=\"{2}\"{3}/>", SecurityElement.Escape(rel.Id), rel.RelationshipType, SecurityElement.Escape(rel.TargetUri.OriginalString), rel.TargetMode == TargetMode.External ? " TargetMode=\"External\"" : "");
                }
            }
            xml.Append("</Relationships>");

            os.PutNextEntry(fileName);
            byte[] b = Encoding.UTF8.GetBytes(xml.ToString());
            os.Write(b, 0, b.Length);
        }

        /// <summary>
        /// Number of items in the collection
        /// </summary>
        public int Count
        {
            get
            {
                return _rels.Count;
            }
        }
    }
}
