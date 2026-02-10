/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  10/07/2025         EPPlus Software AB       Initial release EPPlus 8.3
 *************************************************************************************************/
using OfficeOpenXml.Utils.FileUtils;
using System.Collections;
using System.Collections.Generic;


namespace OfficeOpenXml.Data.CustomXml
{
    /// <summary>
    /// Represents a collection of custom XML parts in the package.
    /// </summary>
    public class ExcelCustomXmlCollection : IEnumerable<ExcelCustomXml>
    {
        ExcelPackage _package;
        List<ExcelCustomXml> _list = new List<ExcelCustomXml>();
        internal ExcelCustomXmlCollection(ExcelPackage package)
        {
            _package = package;
            foreach (var rel in package.Workbook.Part.GetRelationshipsByType("http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml"))
            {
                var cxPart = package.ZipPackage.GetPart(UriHelper.ResolvePartUri(rel.SourceUri, rel.TargetUri));
                var item = new ExcelCustomXml(cxPart);
                _list.Add(item);
            }
        }

        /// <summary>
        /// The indexer for the collection
        /// </summary>
        /// <param name="index">The index</param>
        /// <returns></returns>
        public ExcelCustomXml this[int index]
        {
            get
            {
                return _list[index];
            }
        }
        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }
        /// <summary>
        /// The enumerator for the collection
        /// </summary>
        /// <returns>The enumerator</returns>
        public IEnumerator<ExcelCustomXml> GetEnumerator()
        {
            return _list.GetEnumerator();
        }
        internal void Add(ExcelCustomXml customXml)
        {
            _list.Add(customXml);
        }
        internal bool Contains(ExcelCustomXml customXml)
        {
            return _list.Contains(customXml);
        }
        internal void Remove(ExcelCustomXml customXml)
        {
            _package.ZipPackage.DeletePart(customXml.Part.Uri);
            _package.ZipPackage.DeletePart(customXml.PropertiesPart.Uri);
            _list.Remove(customXml);
        }
        internal void Save()
        {
            foreach(var item in _list)
            {
                item.Save(_package);
            }
        }

        /// <summary>
        /// Number of items in the collection.
        /// </summary>
        public int Count
        {
            get
            {
                return _list.Count;
            }
        }
    }
}
