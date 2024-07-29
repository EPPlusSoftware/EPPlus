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
using System.Xml;

namespace OfficeOpenXml.Table.PivotTable
{
    /// <summary>
    /// A collection of pivottable objects
    /// </summary>
    public class ExcelPivotTableCollection : IEnumerable<ExcelPivotTable>
    {
        List<ExcelPivotTable> _pivotTables = new List<ExcelPivotTable>();
        internal Dictionary<string, int> _pivotTableNames = new Dictionary<string, int>();
        ExcelWorksheet _ws;
        
        internal ExcelPivotTableCollection()
        {
        }
        internal ExcelPivotTableCollection(ExcelWorksheet ws)
        {
            var pck = ws._package.ZipPackage;
            _ws = ws;            
            foreach(var rel in ws.Part.GetRelationships())
            {
                if (rel.RelationshipType == ExcelPackage.schemaRelationships + "/pivotTable")
                {
                    var tbl = new ExcelPivotTable(rel, ws);
                    _pivotTableNames.Add(tbl.Name, _pivotTables.Count);
                    _pivotTables.Add(tbl);
                }
            }
        }
        internal ExcelPivotTable Add(ExcelPivotTable tbl)
        {
            _pivotTables.Add(tbl);
            _pivotTableNames.Add(tbl.Name, _pivotTables.Count - 1);
            if (tbl.CacheId >= _ws.Workbook._nextPivotTableID)
            {
                _ws.Workbook._nextPivotTableID = tbl.CacheId + 1;
            }
            return tbl;
        }

        /// <summary>
        /// Create a pivottable on the supplied range
        /// </summary>
        /// <param name="Range">The range address including header and total row</param>
        /// <param name="Source">The Source data range address</param>
        /// <param name="Name">The name of the pivottable. Must be unique </param>
        /// <returns>The pivottable object</returns>
        public ExcelPivotTable Add(ExcelAddressBase Range, ExcelRangeBase Source, string Name)
        {
            if (string.IsNullOrEmpty(Name))
            {
                Name = GetNewTableName();
            }
            ValidateAdd(Range, Source, Name);
            return Add(new ExcelPivotTable(_ws, Range, Source, Name, _ws.Workbook._nextPivotTableID++));
        }

        private void ValidateAdd(ExcelAddressBase Range, ExcelRangeBase Source, string Name)
        {
            if (Source.Rows < 2)
            {
                throw (new ArgumentException("The Range must contain at least 2 rows", "Source"));
            }
            if (Range.WorkSheetName != _ws.Name)
            {
                throw (new Exception("The Range must be in the current worksheet"));
            }
            else if (_ws.Workbook.ExistsTableName(Name))
            {
                throw (new ArgumentException("Tablename is not unique"));
            }
            foreach (var t in _pivotTables)
            {
                if (t.Address.Collide(Range) != ExcelAddressBase.eAddressCollition.No)
                {
                    throw (new ArgumentException(string.Format("Table range collides with table {0}", t.Name)));
                }
            }
            for (int i = 0; i < Source.Columns; i++)
            {
                if (Source.Offset(0, i, 1, 1).Value == null)
                {
                    throw (new ArgumentException("First row of source range should contain the field headers and can't have blank cells.", "Source"));
                }
            }
        }

        /// <summary>
        /// Create a pivottable on the supplied range
        /// </summary>
        /// <param name="Range">The range address including header and total row</param>
        /// <param name="Source">The source table</param>
        /// <param name="Name">The name of the pivottable. Must be unique </param>
        /// <returns>The pivottable object</returns>
        public ExcelPivotTable Add(ExcelAddressBase Range, ExcelTable Source, string Name)
        {
            if(Source.WorkSheet.Workbook!=_ws.Workbook)
            {
                throw new ArgumentException("The table must be in the same package as the pivottable", "Source");
            }
            if (string.IsNullOrEmpty(Name))
            {
                Name = GetNewTableName();
            }
            ValidateAdd(Range, Source.Range, Name);
            return Add(new ExcelPivotTable(_ws, Range, Source.Range, Name, _ws.Workbook._nextPivotTableID++));
        }
        /// <summary>
        /// Create a pivottable on the supplied range
        /// </summary>
        /// <param name="Range">The range address including header and total row</param>
        /// <param name="PivotCacheDefinition">A pivot table cache shared with another pivot table</param>
        /// <param name="Name">The name of the pivottable. Must be unique </param>
        /// <returns>The pivottable object</returns>
        public ExcelPivotTable Add(ExcelAddressBase Range, ExcelPivotCacheDefinition PivotCacheDefinition, string Name)
        {
            return Add(new ExcelPivotTable(_ws, Range, PivotCacheDefinition._cacheReference, Name, _ws.Workbook._nextPivotTableID++));
        }
        internal string GetNewTableName()
        {
            string name = "Pivottable1";
            int i = 2;
            while (_ws.Workbook.ExistsPivotTableName(name))
            {
                name = string.Format("Pivottable{0}", i++);
            }
            return name;
        }
        /// <summary>
        /// Number of items in the collection
        /// </summary>
        public int Count
        {
            get
            {
                return _pivotTables.Count;
            }
        }
        /// <summary>
        /// The pivottable Index. Base 0.
        /// </summary>
        /// <param name="Index"></param>
        /// <returns></returns>
        public ExcelPivotTable this[int Index]
        {
            get
            {
                if (Index < 0 || Index >= _pivotTables.Count)
                {
                    throw (new ArgumentOutOfRangeException("PivotTable index out of range"));
                }
                return _pivotTables[Index];
            }
        }
        /// <summary>
        /// Pivottabes accesed by name
        /// </summary>
        /// <param name="Name">The name of the pivottable</param>
        /// <returns>The Pivotable. Null if the no match is found</returns>
        public ExcelPivotTable this[string Name]
        {
            get
            {
                if (_pivotTableNames.ContainsKey(Name))
                {
                    return _pivotTables[_pivotTableNames[Name]];
                }
                else
                {
                    return null;
                }
            }
        }
        /// <summary>
        /// Gets the enumerator of the collection
        /// </summary>
        /// <returns>The enumerator</returns>
        public IEnumerator<ExcelPivotTable> GetEnumerator()
        {
            return _pivotTables.GetEnumerator();
        }

        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return _pivotTables.GetEnumerator();
        }
        /// <summary>
        /// Delete the pivottable with the supplied name
        /// </summary>
        /// <param name="Name">The name of the pivottable</param>
        /// <param name="ClearRange">Clear the table range</param>
        public void Delete(string Name, bool ClearRange=false)
            {
            if (!_pivotTableNames.ContainsKey(Name))
            {
                throw new InvalidOperationException($"No pivottable with the name: {Name}");
            }
            Delete(_pivotTables[_pivotTableNames[Name]], ClearRange);
        }
        /// <summary>
        /// Delete the pivot table at the specified index
        /// </summary>
        /// <param name="Index">The index in the PivotTable collection</param>
        /// <param name="ClearRange">Clear the table range</param>
        public void Delete(int Index, bool ClearRange = false)
        {
            if(Index >=0 && Index <_pivotTables.Count)
            {
                throw new IndexOutOfRangeException();
            }
            Delete(_pivotTables[Index], ClearRange);
        }
        /// <summary>
        /// Delete the supplied pivot table 
        /// </summary>
        /// <param name="PivotTable">The PivotTable to remove from the collection</param>
        /// <param name="ClearRange">Clear the table range</param>
        public void Delete(ExcelPivotTable PivotTable, bool ClearRange = false)
        {
            if(!_pivotTables.Contains(PivotTable))
            {
                throw new ArgumentException("This pivot table does not exist in the collection");
            }
            var pck = _ws._package.ZipPackage;

            var cacheReference = PivotTable.CacheDefinition._cacheReference;

            cacheReference._pivotTables.Remove(PivotTable);

            if (cacheReference._pivotTables.Count == 0)
            {
                var cacheAddress = cacheReference.GetSourceAddress();
                var pivotCache = _ws.Workbook._pivotTableCaches[cacheAddress];
                pivotCache.PivotCaches.Remove(cacheReference);
                PivotTable.CacheDefinition._cacheReference.Delete();

                if (pivotCache.PivotCaches.Count == 0)
                {
                    _ws.Workbook._pivotTableIds.Remove(cacheReference.CacheDefinitionUri);
                    _ws.Workbook._pivotTableCaches.Remove(cacheAddress);
                }
            }

            pck.DeletePart(PivotTable.Part.Uri);

            _pivotTables.Remove(PivotTable);
            _pivotTableNames.Remove(PivotTable.Name);
        }
        /// <summary>
        /// Calculate all pivot tables in the collection.
        /// Also see <seealso cref="ExcelPivotTable.Calculate(bool)"/> and <seealso cref="ExcelWorkbook.CalculateAllPivotTables(bool)"/>
        /// </summary>
        /// <param name="refresh">If the cache should be refreshed.</param>
        public void Calculate(bool refresh = false)
        {
            var caches = new HashSet<PivotTableCacheInternal>();
            foreach (var pt in _pivotTables)
            {
                var cache = pt.CacheDefinition._cacheReference;
                if (cache == null) continue;
                if (!caches.Contains(cache))
                {
                    pt.Calculate(refresh);
                    caches.Add(cache);
                }
                else
                {
                    pt.Calculate(false);
                }
            }
        }
    }
}
