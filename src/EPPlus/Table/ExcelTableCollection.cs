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
using System.Xml;
using OfficeOpenXml.Core.RangeQuadTree;
using OfficeOpenXml.Data.Connection;
using OfficeOpenXml.Data.Connection.IOHandlers;
using OfficeOpenXml.Data.QueryTable;
using OfficeOpenXml.FormulaParsing.ExcelUtilities;
namespace OfficeOpenXml.Table
{
    /// <summary>
    /// A collection of table objects
    /// </summary>
    public class ExcelTableCollection : IEnumerable<ExcelTable>
    {
        List<ExcelTable> _tables = new List<ExcelTable>();
        internal Dictionary<string, int> _tableNames = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        ExcelWorksheet _ws;
        internal QuadTree<ExcelTable> qtTables;

        internal ExcelTableCollection(ExcelWorksheet ws)
        {
            var pck = ws._package.ZipPackage;
            _ws = ws;

            if (_ws == null || _ws.Dimension == null)
            {
                qtTables = new QuadTree<ExcelTable>();
            }
            else
            {
                qtTables = new QuadTree<ExcelTable>(_ws.Dimension);
            }

            foreach (XmlElement node in ws.WorksheetXml.SelectNodes("//d:tableParts/d:tablePart", ws.NameSpaceManager))
            {
                var rel = ws.Part.GetRelationship(node.GetAttribute("id",ExcelPackage.schemaRelationships));
                var tbl = new ExcelTable(rel, ws);
                _tableNames.Add(tbl.Name, _tables.Count);
                _tables.Add(tbl);
                qtTables.Add(new QuadRange(tbl.Address), tbl);
            }
        }
        private ExcelTable Add(ExcelTable tbl)
        {
            _tables.Add(tbl);
            _tableNames.Add(tbl.Name, _tables.Count - 1);
            if (tbl.Id >= _ws.Workbook._nextTableID)
            {
                _ws.Workbook._nextTableID = tbl.Id + 1;
            }
            return tbl;
        }

        /// <summary>
        /// Create a table on the supplied range
        /// </summary>
        /// <param name="Range">The range address including header and total row</param>
        /// <param name="Name">The name of the table. Must be unique </param>
        /// <returns>The table object</returns>
        public ExcelTable Add(ExcelAddressBase Range, string Name)
        {
            return AddInternal(Range, Name, null);
        }
        /// <summary>
        /// Adds a new query table to the collection. To add a new query table you need to specify the queries field names. 
        /// </summary>
        /// <param name="Range">The range to the table.</param>
        /// <param name="Name">The name of the table</param>
        /// <param name="connection">The connection in the <see cref="ExcelWorkbook.Connections" collection./> </param>
        /// <param name="fields">The names of the fields. EPPlus does not execute the query, so you must set the field names and update the field properties in the <see cref="ExcelQueryTable.Fields" collection. /></param>
        /// <returns></returns>
        /// <exception cref="ArgumentException">If no fields are supplied</exception>
        /// <exception cref="InvalidOperationException">If the connection is null or not in the package.</exception>
        public ExcelTable AddQueryTable(ExcelAddressBase Range, string Name, ExcelConnection connection, params string[] fields)
        {
            if(fields==null || fields.Length==0 || !fields.Any(x=>!string.IsNullOrEmpty(x)))
            {
                throw new ArgumentException("Supply at least one non-empty field for the query.");
            }
            
            if(connection == null || !_ws.Workbook.Connections.Contains(connection))
            {
                throw new InvalidOperationException("The connection is null or does not exist in the package.");
            }
            
            if(Range.Columns!= fields.Length)
            {
                throw new ArgumentException("The number of fields must match the number of columns in the range.");
            }

            if(connection.Type == eConnectionDataSourceType.Text ||
               connection.Type == eConnectionDataSourceType.WebQuery)
            {
                throw new InvalidOperationException("Legacy web/text connections are not supported for table querys. This type of connection can only be used on worksheet related query tables. Please use the Worksheet.QueryTables.Add method for legacy queries or use a oledb power query connection.");
            }

            var tbl = AddInternal(Range, Name, null);
            tbl.DataSourceType = TableDataSourceType.QueryTable;
            int col = Range.Start.Column;

            var qt = new ExcelQueryTable(new QueryTableDataPartXmlHandler(tbl, connection, fields));
            qt.ConnectionId = connection.Id;
            qt.Connection = connection;
            qt.Name = connection.Name;
            int ix = 1, tfIx=1;
            //Set Column headers.
            foreach (var f in fields)
            {
                if(string.IsNullOrEmpty(f))
                {
                    _ws.Cells[Range.Start.Row, col++].Value = $"Column {tfIx}";
                }
                else
                {
                    _ws.Cells[Range.Start.Row, col++].Value = f;
                    //qt.Fields.Add(new ExcelQueryTableField() { Id=ix, Name=f, TableColumnId=tfIx });
                }                
                tfIx++;
            }
            tbl.QueryTable = qt;

            if (tbl.WorkSheet.Names.ContainsKey(tbl.QueryTable.Name))
            {
                tbl.WorkSheet.Names.Remove(tbl.QueryTable.Name);
            }
            var qtName = tbl.WorkSheet.Names.Add(tbl.QueryTable.Name, tbl.Range);
            qtName.IsNameHidden = true;

            return tbl;
        }
        internal ExcelTable AddInternal(ExcelAddressBase Range, string name, ExcelTable copy)
        {
            if (Range.WorkSheetName != null && Range.WorkSheetName != _ws.Name)
            {
                throw new ArgumentException("Range does not belong to a worksheet", "Range");
            }

            if (string.IsNullOrEmpty(name))
            {
                name = GetNewTableName();
            }
            else
            {
                if (_ws.Workbook.ExistsTableName(name))
                {
                    throw (new ArgumentException("Tablename is not unique"));
                }
            }

            ValidateName(name);

            foreach (var t in _tables)
            {
                if (t.Address.Collide(Range) != ExcelAddressBase.eAddressCollition.No)
                {
                    throw (new ArgumentException(string.Format("Table range collides with table {0}", t.Name)));
                }
            }

            foreach (var mc in _ws.MergedCells)
            {
                if (mc == null) continue; // Issue 780: this happens if a merged cell has been removed
                if (new ExcelAddressBase(mc).Collide(Range) != ExcelAddressBase.eAddressCollition.No)
                {
                    throw (new ArgumentException($"Table range collides with merged range {mc}"));
                }
            }

            var tableToAdd = new ExcelTable(_ws, Range, name, _ws.Workbook._nextTableID, copy);
            qtTables.Add(new QuadRange(tableToAdd.Address), tableToAdd);
            return Add(tableToAdd);
        }

        private void ValidateName(string name)
        {
            if (string.IsNullOrEmpty(name.Trim()))
            {
                throw new ArgumentException("Tablename is blank", "Name");
            }

            var c = name[0];
            if (char.IsLetter(c) == false && c != '\\' && c != '_')
            {
                throw new ArgumentException("Tablename start with invalid character", "Name");
            }

            if (!ExcelAddressUtil.IsValidName(name))
            {
                throw (new ArgumentException("Tablename is not valid", "Name"));
            }
        }
        /// <summary>
        /// Delete the table at the specified index
        /// </summary>
        /// <param name="Index">The index</param>
        /// <param name="ClearRange">Clear the rage if set to true</param>
        public void Delete(int Index, bool ClearRange = false)
        {
            Delete(this[Index], ClearRange);
        }

        /// <summary>
        /// Delete the table with the specified name
        /// </summary>
        /// <param name="Name">The name of the table to be deleted</param>
        /// <param name="ClearRange">Clear the rage if set to true</param>
        public void Delete(string Name, bool ClearRange = false)
        {
            if (this[Name] == null)
            {
                throw new ArgumentOutOfRangeException(string.Format("Cannot delete non-existant table {0} in sheet {1}.", Name, _ws.Name));
            }
            Delete(this[Name], ClearRange);
        }


        /// <summary>
        /// Delete the table and removes it from the collection
        /// </summary>
        /// <param name="Table">The table object</param>
        /// <param name="ClearRange">Clear the table range</param>
        public void Delete(ExcelTable Table, bool ClearRange = false)
        {
            if (!this._tables.Contains(Table))
            {
                throw new ArgumentOutOfRangeException("Table", String.Format("Table {0} does not exist in this collection", Table.Name));
            }
            lock (this)
            {
                var tIx = _tableNames[Table.Name];
                _tableNames.Remove(Table.Name);
                _tables.Remove(Table);
                foreach (var sheet in Table.WorkSheet.Workbook.Worksheets)
                {
                    if (sheet is ExcelChartsheet) continue;
                    foreach (var t in sheet.Tables)
                    {
                        if (t.Id > Table.Id) t.Id--;
                    }
                }
                foreach(var name in _tableNames.Keys.ToArray())
                { 
                    if(_tableNames[name] > tIx)
                    {
                        _tableNames[name]--;
                    }
                }
                Table.DeleteMe();
                if (ClearRange)
                {
                    var range = _ws.Cells[Table.Address.Address];
                    range.Clear();
                }
                if(Table.DataSourceType==TableDataSourceType.QueryTable && Table.QueryTable!=null)
                {
                    var qtName = Table.WorkSheet.Names[Table.QueryTable.Name];
                    if (qtName != null)
                    {
                        Table.WorkSheet.Names.Remove(qtName.Name);
                    }
                    Table.QueryTable.Remove();
                }
            }
        }

        internal string GetNewTableName(string name = "Table{0}")
        {            
            var newName = string.Format(name, 1);
            if (newName == name) name += "{0}";
            int i = 2;
            while (_ws.Workbook.ExistsTableName(newName))
            {
                newName = string.Format(name, i++);
            }
            return newName;
        }
        /// <summary>
        /// Number of items in the collection
        /// </summary>
        public int Count
        {
            get
            {
                return _tables.Count;
            }
        }
        /// <summary>
        /// Get the table object from a range.
        /// </summary>
        /// <param name="Range">The range</param>
        /// <returns>The table. Null if no range matches</returns>
        public ExcelTable GetFromRange(ExcelRangeBase Range)
        {
            foreach (var tbl in Range.Worksheet.Tables)
            {
                if (tbl.Address._address == Range._address)
                {
                    return tbl;
                }
            }
            return null;
        }
        /// <summary>
        /// The table Index. Base 0.
        /// </summary>
        /// <param name="Index"></param>
        /// <returns></returns>
        public ExcelTable this[int Index]
        {
            get
            {
                if (Index < 0 || Index >= _tables.Count)
                {
                    throw (new ArgumentOutOfRangeException("Table index out of range"));
                }
                return _tables[Index];
            }
        }
        /// <summary>
        /// Indexer
        /// </summary>
        /// <param name="Name">The name of the table</param>
        /// <returns>The table. Null if the table name is not found in the collection</returns>
        public ExcelTable this[string Name]
        {
            get
            {
                if (_tableNames.ContainsKey(Name))
                {
                    return _tables[_tableNames[Name]];
                }
                else
                {
                    return null;
                }
            }
        }
        /// <summary>
        /// Gets the enumerator for the collection
        /// </summary>
        /// <returns>The enumerator</returns>
        public IEnumerator<ExcelTable> GetEnumerator()
        {
            return _tables.GetEnumerator();
        }

        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return _tables.GetEnumerator();
        }

        internal List<QuadRangeItem<ExcelTable>> GetIntersectingRanges(ExcelAddress address)
        {
            return qtTables.GetIntersectingRangeItems(new QuadRange(address));
        }
    }
}
