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
using OfficeOpenXml.Constants;
using OfficeOpenXml.Packaging;
using OfficeOpenXml.Utils.FileUtils;
using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml;

namespace OfficeOpenXml.Data.Connection
{
    /// <summary>
    /// A collection of external connections in a workbook.
    /// </summary>
    public class ExcelConnectionCollection : IEnumerable<ExcelConnection>
    {
        ExcelPackage _package;
        XmlNamespaceManager _nsm;
        List<ExcelConnection> _list = new List<ExcelConnection>();
        int _nextId=1;
        internal  ExcelConnectionCollection(ExcelPackage package)
        {
            _package = package;
            Part = _package.ZipPackage.GetByContentType(ContentTypes.contentTypeConnections).FirstOrDefault();
            _nsm = _package.Workbook.NameSpaceManager;
            if (Part!=null)
            {
                ConnectionXml = new XmlDocument();
                XmlHelper.LoadXmlSafe(ConnectionXml, Part.GetStream());
                foreach (XmlNode node in ConnectionXml.DocumentElement.SelectNodes("d:connection", _nsm))
                {
                    var c = new ExcelConnection(new ConnectionDataPartXmlHandler(_nsm, node));
                    if(c.Id>=_nextId) _nextId= c.Id+1;
                    _list.Add(c);
                }
            }
        }
        internal ZipPackagePart Part { get; private set; }
        internal XmlDocument ConnectionXml { get; private set; }
        /// <summary>
        /// Number of items in the collection.
        /// </summary>
        public int Count { get { return _list.Count; } }
        public IEnumerator<ExcelConnection> GetEnumerator()
        {
            return _list.GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return _list.GetEnumerator();
        }
        /// <summary>
        /// Returns the connection at the supplied position.
        /// </summary>
        /// <param name="index">The index of the connection to return.</param>
        /// <returns>The connection</returns>
        public ExcelConnection this[int index]
        {
            get
            {
                return _list[index];
            }
        }

        internal void Save()
        {
            if (_list.Count==0 && Part!=null)
            {
                _package.ZipPackage.DeletePart(Part.Uri);
                Part = null;
                ConnectionXml = null;
                return;
            }
            if(Part==null)
            {
                Part = _package.ZipPackage.CreatePart(new Uri("/xl/connections.xml", UriKind.Relative), ContentTypes.contentTypeConnections, CompressionLevel.Default);
                _package.Workbook.Part.CreateRelationship(Part.Uri, TargetMode.Internal, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/connections");
                ConnectionXml = new XmlDocument();
                XmlHelper.LoadXmlSafe(ConnectionXml, "<?xml version=\"1.0\" encoding=\"utf-8\"?><connections xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"></connections>", System.Text.Encoding.UTF8);
            }
            foreach (var conn in _list)
            {
                conn.Save();
            }
            var connXmlStream = Part.GetStream(System.IO.FileMode.Create, System.IO.FileAccess.Write);
            ConnectionXml.Save(connXmlStream);
            connXmlStream.Flush();
        }
        /// <summary>
        /// Adds a connection of type database with the specified connection string. 
        /// EPPlus will set the <see cref="Type"/> of the connection to Odbc or OleDb, depending on the connection string. If you use another type, please set it manually.
        /// If the connection string is a power querty connection string, please also see <see cref="ExcelWorkbook.PowerQuerySettings"/>
        /// </summary>-
        /// <param name="name">The name of the connection.</param>
        /// <param name="connectionString">The connection string to the database.</param>
        /// <returns>The connection</returns>
        public ExcelConnection AddDatabase(string name, string connectionString)
        {
            if (string.IsNullOrEmpty(connectionString?.Trim()))
            {
                throw new ArgumentException("Connection string cannot be null or empty", nameof(connectionString));
            };

            ExcelConnection c = AddInternal(name);
            c.DatabaseProperties = new ExcelDatabaseProperties();
            c.DatabaseProperties.Connection = connectionString;
            c.Type = GetConnectionType(connectionString);
            c.IsBackground = true;
            c.SaveData = true;
            
            return c;
        }
        /// <summary>
        /// Adds a power query connection with the specified connection string. 
        /// EPPlus will set the <see cref="Type"/> of the connection to OleDb.
        /// A power query connection requires a connection string with the OleDb provide Microsoft.Mashup.OleDb.1 and a M-Formula containing the query.
        /// The <see cref="ExcelWorkbook.PowerQuerySettings"/> property will be created, if it does not exist.
        /// </summary>
        /// <param name="name">The name of the connection.</param>
        /// <param name="connectionString">The connection string to the database. Power query connection string usually uses the Microsoft.Mashup.OleDb.1 OleDb provider. For example: Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location="Table 1";Extended Properties=""</param>
        /// <param name="mFormula">The M formulas to use for the power query connection, without the Section1 declaration. This formula is appended to the Formulas property of the <see cref="ExcelWorkbook.PowerQuerySettings"/> object.</param>
        /// <returns>The connection</returns>
        public ExcelConnection AddPowerQuery(string name, string connectionString, string mFormula)
        {
            if(mFormula==null || mFormula.Trim().Length==0)
            {
                throw new ArgumentException("M-Formula cannot be null or empty", nameof(mFormula));
            }
            if(mFormula.StartsWith("section Section1",StringComparison.InvariantCultureIgnoreCase))
            {
                throw new ArgumentException("The M-Formula should not contain the Section1 declaration.", nameof(mFormula));
            }
            var c = AddDatabase(name, connectionString);
            if (_package.Workbook.PowerQuerySettings.Exists == false)
            {
                _package.Workbook.PowerQuerySettings.Create();
            }
            if(_package.Workbook.PowerQuerySettings.Formulas.EndsWith("\r") ||
               _package.Workbook.PowerQuerySettings.Formulas.EndsWith("\n"))
            {
                _package.Workbook.PowerQuerySettings.Formulas += "\r\n";
            }
            _package.Workbook.PowerQuerySettings.Formulas += mFormula;
            return c;
        }
        private static eConnectionDataSourceType GetConnectionType(string connectionString)
        {
            var parameters = connectionString.Split(';');
            foreach (var param in parameters)
            {
                var keyValue = param.Split('=');
                switch (keyValue[0]?.Trim().ToLower())
                {
                    case "driver":
                        return eConnectionDataSourceType.ODBC;
                    case "provider":
                        return eConnectionDataSourceType.OLEDB;
                }
            }
            return eConnectionDataSourceType.ODBC;
        }

        /// <summary>
        /// Adds a connection to a OLAP data source.
        /// </summary>
        /// <param name="name">The uniqe name of the connection</param>
        /// <param name="connectionString">The connection string </param>
        /// <param name="command">The command, usually the name of the qube.</param>
        /// <returns>The connection</returns>
        public ExcelConnection AddOlap(string name, string connectionString, string command)
        {
            if (string.IsNullOrEmpty(connectionString?.Trim()))
            {
                throw new ArgumentException("Connection string cannot be null or empty", nameof(connectionString));
            }

            ExcelConnection c = AddInternal(name);
            c.DatabaseProperties = new ExcelDatabaseProperties()
            {
                Connection = connectionString,
                CommandType = eCommandType.Cube,
                Command = command
            };
            c.Type = GetConnectionType(connectionString);
            c.OlapProperties = new ExcelConnectionOlapProperties();
            return c;
        }
        /// <summary>
        /// Adds a connection to a web query data source.
        /// </summary>
        /// <returns>The connection</returns>
        public ExcelConnection AddWeb(string name, string url)
        {
            if (string.IsNullOrEmpty(url?.Trim()))
            {
                throw new ArgumentException("Connection string cannot be null or empty", nameof(url));
            }

            ExcelConnection c = AddInternal(name);
            c.WebProperties = new ExcelWebProperties();
            c.WebProperties.Url = url;
            c.Type = eConnectionDataSourceType.WebQuery;
            return c;
        }
        /// <summary>
        /// Adds a connection to a text file data source.
        /// </summary>
        /// <param name="name">The name of the connection</param>
        /// <param name="sourceFile">The path to the text file to use to import external data. Can be expressed in URI or system specific file path notation.</param>
        /// <returns>The connection</returns>
        public ExcelConnection AddText(string name, FileInfo sourceFile)
        {
            return AddText(name, sourceFile.FullName);
        }
        /// <summary>
        /// Adds a connection to a text file data source.
        /// </summary>
        /// <param name="name">The name of the connection</param>
        /// <param name="sourceFile">The path to the text file to use to import external data. Can be expressed in URI or system specific file path notation.</param>
        /// <returns>The connection</returns>
        public ExcelConnection AddText(string name, string sourceFile)
        {
            if(string.IsNullOrEmpty(sourceFile?.Trim()))
            {
                throw new ArgumentException($"Argument {nameof(sourceFile)} cannot be empty.");
            }
            ExcelConnection c = AddInternal(name);
            c.TextProperties = new ExcelTextProperties();
            c.TextProperties.SourceFile=sourceFile;
            c.Type = eConnectionDataSourceType.Text;
            return c;
        }

        private ExcelConnection AddInternal(string name)
        {
            if(Part==null)
            {
                CreatePartAndXml();
            }
            else
            {
                if (_list.Any(x=>x.Name.Equals(name, StringComparison.CurrentCultureIgnoreCase)))
                {
                    throw new ArgumentException("A connection with name {name} already exist in the collection.");
                }
            }
            
            var node = ConnectionXml.CreateElement("connection", Schemas.schemaMain);
            ConnectionXml.DocumentElement.AppendChild(node);
            node.SetAttribute("uid", "http://schemas.microsoft.com/office/spreadsheetml/2017/revision16", "{" + Guid.NewGuid().ToString().ToUpperInvariant() + "}");
            var c = new ExcelConnection(new ConnectionDataPartXmlHandler(_nsm, node));
            c.Id = _nextId++;
            c.Name = name;
            c.LastRefreshVersion = 8;
            c.Parameters = new ExcelConnectionParameters();
            _list.Add(c);
            return c;
        }

        private void CreatePartAndXml()
        {
            var uri = new Uri("/xl/connections.xml", UriKind.Relative);
            Part = _package.ZipPackage.CreatePart(uri, ContentTypes.contentTypeConnections);
            var rel = _package.Workbook.Part.CreateRelationship(UriHelper.GetRelativeUri(_package.Workbook.WorkbookUri, uri), TargetMode.Internal, ExcelPackage.schemaRelationships + "/connections");
            ConnectionXml = new XmlDocument();
            var startXml = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><connections xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" mc:Ignorable=\"xr16\" xmlns:xr16=\"http://schemas.microsoft.com/office/spreadsheetml/2017/revision16\"></connections>";
            ConnectionXml.LoadXml(startXml);
        }
        /// <summary>
        /// Removes the connection at the given position. Please note that any related Power Query formula is not removed for the <see cref="ExcelPowerQuerySettings.Formulas"/> by this method.
        /// </summary>
        /// <param name="index">The position of the connection to remove.</param>
        public void RemoveAt(int index)
        {
            _list[index].Remove();
            _list.RemoveAt(index);
        }
        /// <summary>
        /// Removes the connection at the given position. Please note that any related Power Query formula is not removed from the <see cref="ExcelPowerQuerySettings.Formulas"/> by this method.
        /// </summary>
        /// <param name="connection">The connection to remove.</param>
        public void Remove(ExcelConnection connection)
        {
            foreach(var ws in _package.Workbook.Worksheets)
            {
                if(ws.QueryTables.Any(x => x.ConnectionId == connection.Id))
                {
                    var qt = ws.QueryTables.First(x => x.ConnectionId == connection.Id);
                    throw new InvalidOperationException($"Can not remove connection with id {connection.Id}. The connection is used by a query table in the worksheet {ws.Name} in range {qt.DestinationRange.Address}.");
                }
                if(ws.Tables.Any(t => t.DataSourceType == TableDataSourceType.QueryTable && t.QueryTable.ConnectionId == connection.Id))
                {
                    var t = ws.Tables.First(t => t.DataSourceType == TableDataSourceType.QueryTable && t.QueryTable.ConnectionId == connection.Id);
                    throw new InvalidOperationException($"Can not remove connection with id {connection.Id}. The connection is used by a table {t.Name} in the worksheet {ws.Name}.");
                }
            }
            connection.Remove();
            _list.Remove(connection);
        }
    }
}
