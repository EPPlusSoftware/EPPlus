using OfficeOpenXml.Constants;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Engineering;
using OfficeOpenXml.Packaging;
using OfficeOpenXml.Utils.FileUtils;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;


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

        internal void Save()
        {
            foreach(var item in _list)
            {
                item.Save();
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
