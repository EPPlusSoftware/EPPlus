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
using OfficeOpenXml.FormulaParsing.Excel.Functions.MathFunctions;
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
        /// If the connection string is a power querty connection string, please also see <see cref="PowerQuerySettings"/>
        /// </summary>
        /// <param name="name">The name of the connection.</param>
        /// <param name="connectionString">The connection string to the database.</param>
        /// <returns>The connection</returns>
        public ExcelConnection AddDatabase(string name, string connectionString)
        {
            if (string.IsNullOrEmpty(connectionString?.Trim()))
            {
                throw new ArgumentException("Connection string cannot be null or empty", nameof(connectionString));
            }
            ;
            ExcelConnection c = AddInternal(name);
            c.DatabaseProperties = new ExcelDatabaseProperties();
            c.DatabaseProperties.Connection = connectionString;
            c.Type = GetConnectionType(connectionString);
            c.IsBackground = true;
            c.SaveData = true;

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
            c.Type = eConnectionDataSourceType.OLEDB;
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
        /// <param name="sourceFile">The path to the text file to use to import external data. Can be expressed in URI or systemspecific file path notation.</param>
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
        public void RemoveAt(int index)
        {
            _list[index].Remove();
            _list.RemoveAt(index);
        }
        public void Remove(ExcelConnection item)
        {
            item.Remove();
            _list.Remove(item);
        }
        ExcelPowerQuerySettings _powerQuerySettings = null;
        /// <summary>
        /// Settings for Power Query connections/queries. These setting is loaded from the custom XML part with the DataMashup schema.
        /// </summary>
        public ExcelPowerQuerySettings PowerQuerySettings
        {
            get
            {
                if (_powerQuerySettings == null)
                {
                    var pqCustomXml = _package.Workbook.CustomXmlDocuments.FirstOrDefault(x => x.SchemasReferences.Any(x => x == Schemas.schemaDataMashup));
                    if (pqCustomXml == null)
                    {
                        _powerQuerySettings = new ExcelPowerQuerySettings();
                    }
                    else
                    {
                        var blob = Convert.FromBase64String(pqCustomXml.CustomXml.DocumentElement.InnerText);
                        _powerQuerySettings = new ExcelPowerQuerySettings(blob);
                    }
                }
                return  _powerQuerySettings;
            }        
        }
    }
}
