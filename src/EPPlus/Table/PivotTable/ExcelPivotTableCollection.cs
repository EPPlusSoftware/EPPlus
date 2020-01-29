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
        internal ExcelPivotTableCollection(ExcelWorksheet ws)
        {
            var pck = ws._package.Package;
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
        private ExcelPivotTable Add(ExcelPivotTable tbl)
        {
            _pivotTables.Add(tbl);
            _pivotTableNames.Add(tbl.Name, _pivotTables.Count - 1);
            if (tbl.CacheID >= _ws.Workbook._nextPivotTableID)
            {
                _ws.Workbook._nextPivotTableID = tbl.CacheID + 1;
            }
            return tbl;
        }

        /// <summary>
        /// Create a pivottable on the supplied range
        /// </summary>
        /// <param name="Range">The range address including header and total row</param>
        /// <param name="Source">The Source data range address</param>
        /// <param name="Name">The name of the table. Must be unique </param>
        /// <returns>The pivottable object</returns>
        public ExcelPivotTable Add(ExcelAddressBase Range, ExcelRangeBase Source, string Name)
        {
            if (string.IsNullOrEmpty(Name))
            {
                Name = GetNewTableName();
            }
            if (Range.WorkSheet != _ws.Name)
            {
                throw(new Exception("The Range must be in the current worksheet"));
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
            
            return Add(new ExcelPivotTable(_ws, Range, Source, Name, _ws.Workbook._nextPivotTableID++));
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
    }
}
