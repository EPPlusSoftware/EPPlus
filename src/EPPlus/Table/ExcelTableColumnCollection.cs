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
using OfficeOpenXml.Core;
using OfficeOpenXml.Encryption;
using OfficeOpenXml.FormulaParsing.Excel.Functions.MathFunctions;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Xml;

namespace OfficeOpenXml.Table
{
    /// <summary>
    /// A collection of table columns
    /// </summary>
    public class ExcelTableColumnCollection : IEnumerable<ExcelTableColumn>
    {
        List<ExcelTableColumn> _cols = new List<ExcelTableColumn>();
        Dictionary<string, int> _colNames = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        internal int _maxId = 1;
        internal ExcelTableColumnCollection(ExcelTable table)
        {
            Table = table;
            foreach(XmlNode node in table.TableXml.SelectNodes("//d:table/d:tableColumns/d:tableColumn",table.NameSpaceManager))
            {
                var item = new ExcelTableColumn(table.NameSpaceManager, node, table, _cols.Count);
                _cols.Add(item);
                _colNames.Add(_cols[_cols.Count - 1].Name, _cols.Count - 1);
                var id = item.Id;
                if (id>=_maxId)
                {
                    _maxId = id+1;
                }
            }
        }
        /// <summary>
        /// A reference to the table object
        /// </summary>
        public ExcelTable Table
        {
            get;
            private set;
        }
        /// <summary>
        /// Number of items in the collection
        /// </summary>
        public int Count
        {
            get
            {
                return _cols.Count;
            }
        }
        /// <summary>
        /// The column Index. Base 0.
        /// </summary>
        /// <param name="Index"></param>
        /// <returns></returns>
        public ExcelTableColumn this[int Index]
        {
            get
            {
                if (Index < 0 || Index >= _cols.Count)
                {
                    throw (new ArgumentOutOfRangeException("Column index out of range"));
                }
                return _cols[Index] as ExcelTableColumn;
            }
        }
        /// <summary>
        /// Indexer
        /// </summary>
        /// <param name="Name">The name of the table</param>
        /// <returns>The table column. Null if the table name is not found in the collection</returns>
        public ExcelTableColumn this[string Name]
        {
            get
            {
                if (_colNames.ContainsKey(Name))
                {
                    return _cols[_colNames[Name]];
                }
                else
                {
                    return null;
                }
            }
        }

        internal void UpdateColName(int dictIndex, string newName)
        {
            var pair = _colNames.FirstOrDefault(t => t.Value == dictIndex);
            if(pair.Key != null)
            {
                _colNames.Remove(pair.Key);
                _colNames.Add(newName, dictIndex);
            }
        }

        IEnumerator<ExcelTableColumn> IEnumerable<ExcelTableColumn>.GetEnumerator()
        {
            return _cols.GetEnumerator();
        }

        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return _cols.GetEnumerator();
        }
        internal string GetUniqueName(string name)
        {            
            if (_colNames.ContainsKey(name))
            {
                var newName = name;
                var i = 2;
                do
                {
                    newName = name+(i++).ToString(CultureInfo.InvariantCulture);
                }
                while (_colNames.ContainsKey(newName));
                return newName;
            }
            return name;
        }
        /// <summary>
        /// Adds one or more columns at the end of the table.
        /// </summary>
        /// <param name="columns">Number of columns to add.</param>
        /// <returns>The added range</returns>
        public ExcelRangeBase Add(int columns = 1)
        {
            return Insert(int.MaxValue, columns);
        }
        /// <summary>
        /// Inserts one or more columns before the specified position in the table.
        /// </summary>
        /// <param name="position">The position in the table where the column will be inserted. 0 will insert the column at the leftmost position. Any value larger than the number of rows in the table will insert a row at the end of the table.</param>
        /// <param name="columns">Number of columns to insert.</param>
        /// <returns>The inserted range</returns>
        public ExcelRangeBase Insert(int position, int columns = 1)
        {
            lock(Table)
            {
                var range = Table.InsertColumn(position, columns);
                XmlNode refNode;
                if (position >= _cols.Count)
                {
                    refNode = _cols[_cols.Count - 1].TopNode;
                    position = _cols.Count;
                }
                else
                {
                    refNode = _cols[position].TopNode;
                }
                for (int i = position; i < position + columns; i++)
                {
                    var node = Table.TableXml.CreateElement("tableColumn", ExcelPackage.schemaMain);

                    if (i >= _cols.Count)
                    {
                        refNode.ParentNode.AppendChild(node);
                    }
                    else
                    {
                        refNode.ParentNode.InsertBefore(node, refNode);
                    }
                    var item = new ExcelTableColumn(Table.NameSpaceManager, node, Table, i);
                    item.Name = GetUniqueName($"Column{i + 1}");
                    item.Id = _maxId++;
                    _cols.Insert(i, item);
                }
                for (int i = position; i < _cols.Count; i++)
                {
                    _cols[i].Position = i;
                }
                _colNames = _cols.ToDictionary(x => x.Name, y => y.Position);
                return range;
            }
        }
        /// <summary>
        /// Deletes one or more columns from the specified position in the table.
        /// </summary>
        /// <param name="position">The position in the table where the column will be inserted. 0 will insert the column at the leftmost position. Any value larger than the number of rows in the table will insert a row at the end of the table.</param>
        /// <param name="columns">Number of columns to insert.</param>
        /// <returns>The inserted range</returns>
        public ExcelRangeBase Delete(int position, int columns = 1)
        {
            lock (Table)
            {
                if (position < 0)
                {
                    throw new ArgumentException("position", "position can't be negative");
                }
                if (columns < 0)
                {
                    throw new ArgumentException("columns", "columns can't be negative");
                }

                if (Table.Address._toCol < Table.Address._fromCol + position + columns - 1)
                {
                    throw new InvalidOperationException("Delete will exceed the number of columns in the table");
                }

                for (int i = position + columns - 1; i >= position; i--)
                {
                    var n = Table.Columns[i].TopNode;
                    n.ParentNode.RemoveChild(n);
                    Table.Columns._colNames.Remove(_cols[i].Name);
                    Table.Columns._cols.RemoveAt(i);
                }

                for (int i = position; i < _cols.Count; i++)
                {
                    _cols[i].Position = i;
                }

                _colNames = _cols.ToDictionary(x => x.Name, y => y.Position);

                var range = Table.DeleteColumn(position, columns);
                return range;
            }
        }

        internal List<string> GetColNamesList()
        {
            return _colNames.Keys.ToList();
        }

        internal bool ContainsColName(string name)
        {
            return _colNames.ContainsKey(name);
        }

        internal void UpdateColName(string oldName, string newName)
        {
            if(oldName.Equals(newName, StringComparison.InvariantCulture)) return;

            if(_colNames.ContainsKey(oldName))
            {
                var columnIndex = _colNames[oldName];
                if (_colNames.ContainsKey(newName))
                {
                    throw new ArgumentException($"Table '{Table.Name}' cannot add column name '{newName}' at column{columnIndex}. " +
                        $"Name already exists at column{_colNames[newName]}. Column names must be unique.");
                }
                _colNames.Remove(oldName);
                _colNames.Add(newName, columnIndex);
            }
        }

        internal int GetIndexOfColName(string name)
        {
            return _colNames[name];
        }
    }
}
