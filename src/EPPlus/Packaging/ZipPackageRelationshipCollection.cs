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
using System.Text;
using System.Security;
using OfficeOpenXml.Packaging.Ionic.Zip;
using OfficeOpenXml.Utils;

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

        internal string savedFileName = null;

        internal void WriteZip(ZipOutputStream os, string fileName)
        {
            os.PutNextEntry(fileName);
            savedFileName = fileName;
            byte[] b = Encoding.UTF8.GetBytes(GetRelsXML());
            os.Write(b, 0, b.Length);
        }

        internal string GetRelsXML()
        {
            StringBuilder xml = new StringBuilder("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">");
            foreach (var rel in _rels.Values)
            {
                if (rel.TargetUri == null || rel.TargetUri.OriginalString.StartsWith("Invalid:URI", StringComparison.OrdinalIgnoreCase))
                {
                    xml.AppendFormat("<Relationship Id=\"{0}\" Type=\"{1}\" Target=\"{2}\"{3}/>", SecurityElement.Escape(rel.Id), rel.RelationshipType, ConvertUtil.CropString(SecurityElement.Escape(rel.Target), 2079), rel.TargetMode == TargetMode.External ? " TargetMode=\"External\"" : "");
                }
                else
                {
                    xml.AppendFormat("<Relationship Id=\"{0}\" Type=\"{1}\" Target=\"{2}\"{3}/>", SecurityElement.Escape(rel.Id), rel.RelationshipType, ConvertUtil.CropString(SecurityElement.Escape(rel.TargetUri.OriginalString), 2079), rel.TargetMode == TargetMode.External ? " TargetMode=\"External\"" : "");
                }
            }
            xml.Append("</Relationships>");

            return xml.ToString();
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
