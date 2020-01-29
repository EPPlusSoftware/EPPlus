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
using System.Globalization;
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
        internal ExcelTableColumnCollection(ExcelTable table)
        {
            Table = table;
            foreach(XmlNode node in table.TableXml.SelectNodes("//d:table/d:tableColumns/d:tableColumn",table.NameSpaceManager))
            {                
                _cols.Add(new ExcelTableColumn(table.NameSpaceManager, node, table, _cols.Count));
                _colNames.Add(_cols[_cols.Count - 1].Name, _cols.Count - 1);
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
    }
}
