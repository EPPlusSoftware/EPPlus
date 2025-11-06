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
using OfficeOpenXml.Constants;
using OfficeOpenXml.Data.Connection;
using OfficeOpenXml.Data.Connection.IOHandlers;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using OfficeOpenXml.Table;
using OfficeOpenXml.FormulaParsing.Excel.Functions.MathFunctions;
namespace OfficeOpenXml.Data.QueryTable
{
    /// <summary>
    /// A collection of legacy query tables in a worksheet.
    /// Also see <see cref="ExcelTableCollection.AddQueryTable(ExcelAddressBase, string, ExcelConnection, string[])"/>
    /// </summary>
    public class ExcelQueryTableCollection : IEnumerable<ExcelQueryTable>
    {
        private ExcelWorksheet _ws;
        List<ExcelQueryTable> _list = new List<ExcelQueryTable>();
        internal ExcelQueryTableCollection(ExcelWorksheet ws)
        {
            _ws = ws;
            _list = new List<ExcelQueryTable>();

            var rels = ws.Part.GetRelationshipsByType(ExcelPackage.schemaRelationships + "/queryTable");
            foreach(var rel in rels)
            {
                var qt = new ExcelQueryTable(new QueryTableDataPartXmlHandler(ws, rel));
                _list.Add(qt);
            }
        }
        /// <summary>
        /// The query table index. 
        /// </summary>
        /// <param name="index">The index in the collection. </param>
        /// <returns>The query table</returns>
        public ExcelQueryTable this[int index]
        {
            get
            {
                if (index < 0 || index >= _list.Count)
                {
                    throw (new ArgumentOutOfRangeException("Table index out of range"));
                }
                return _list[index];
            }
        }
        /// <summary>
        /// Indexer
        /// </summary>
        /// <param name="name">The name of the table</param>
        /// <returns>The table. Null if the table name is not found in the collection</returns>
        public ExcelQueryTable this[string name]
        {
            get
            {
                return _list.FirstOrDefault(x => x.Name.Equals(name));
            }
        }
        /// <summary>
        /// Adds a new query table to the collection.
        /// </summary>
        /// <param name="address">The address</param>
        /// <param name="name">The name of the query table</param>
        /// <param name="connection">The connection </param>
        /// <returns></returns>
        /// <exception cref="ArgumentException"></exception>
        public ExcelQueryTable Add(ExcelAddressBase address, string name, ExcelConnection connection)
        {
            if(_ws.Workbook.Connections.Contains(connection)==false)
            {
                throw new ArgumentException("The connection must be from the same workbook.", nameof(connection));
            }
            if (string.IsNullOrEmpty(name))
            {
                throw new ArgumentException("Name cannot be null or empty", nameof(name));
            }
            if(_list.Any(x=> x.Name.Equals(name, StringComparison.OrdinalIgnoreCase)))
            {
                throw new ArgumentException($"A query table with the name '{name}' already exists in the collection", nameof(name));
            }

            var qt = new ExcelQueryTable(new QueryTableDataPartXmlHandler(_ws));
            qt.Name = name;
            qt.ConnectionId = connection.Id;
            qt.Connection = connection;
            
            var definedName = _ws.Names.Add(qt.Name, _ws.Cells[address.Address]);
            qt.DestinationRange = definedName;
            _list.Add(qt);
            return qt;
        }
        /// <summary>
        /// The number of query tables in the collection
        /// </summary>
        public int Count 
        {
            get
            {
                return _list.Count;
            }
        }
        /// <summary>
        /// The enumerator for the collection
        /// </summary>
        /// <returns></returns>
        public IEnumerator<ExcelQueryTable> GetEnumerator()
        {
            return _list.GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }
        internal void Save()
        {
            foreach(var qt in _list)
            {
                qt.Save();
            }
        }
    }
}