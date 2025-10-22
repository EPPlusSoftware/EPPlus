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
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Reflection.Emit;
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
        internal  ExcelConnectionCollection(ExcelPackage package)
        {
            _package = package;
            Part = _package.ZipPackage.GetByContentType(ContentTypes.contentTypeConnections).FirstOrDefault();
            ConnectionXml = new XmlDocument();
            _nsm = _package.Workbook.NameSpaceManager;
            if (Part!=null)
            {
                XmlHelper.LoadXmlSafe(ConnectionXml, Part.GetStream());
                foreach (XmlNode node in ConnectionXml.DocumentElement.SelectNodes("d:connection", _nsm))
                {
                    _list.Add(new ExcelConnection(new ConnectionDataPartXmlHandler(_nsm, node)));
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
            if(_list.Count==0 && Part!=null)
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
        /// If the connection string is a power querty connection string, please also see <see cref="PowerQuerySettings"/>
        /// </summary>
        /// <param name="connectionString">The connection string to the database.</param>
        /// <returns>The connection</returns>
        public ExcelConnection AddDatabase(string connectionString)
        {
            if(string.IsNullOrEmpty(connectionString?.Trim()))
            {
                throw new ArgumentException("Connection string cannot be null or empty", nameof(connectionString));
            };
            ExcelConnection c = AddInternal();
            c.DatabaseProperties = new ExcelDatabaseProperties();
            c.DatabaseProperties.Connection = connectionString;
            return c;
        }
        /// <summary>
        /// Adds a connection to a OLAP data source.
        /// </summary>
        /// <returns>The connection</returns>
        public ExcelConnection AddOlap()
        {
            ExcelConnection c = AddInternal();
            c.OlapProperties = new ExcelConnectionOlapProperties();
            return c;
        }
        /// <summary>
        /// Adds a connection to a web query data source.
        /// </summary>
        /// <returns>The connection</returns>
        public ExcelConnection AddWeb()
        {
            ExcelConnection c = AddInternal();
            c.WebProperties = new ExcelWebProperties();
            return c;
        }
        /// <summary>
        /// Adds a connection to a text file data source.
        /// </summary>
        /// <returns>The connection</returns>
        public ExcelConnection AddText()
        {
            ExcelConnection c = AddInternal();
            c.TextProperties = new ExcelTextProperties();
            return c;
        }

        private ExcelConnection AddInternal()
        {
            var node = ConnectionXml.CreateElement("d", "connection", Schemas.schemaMain);
            ConnectionXml.DocumentElement.AppendChild(node);
            var c = new ExcelConnection(new ConnectionDataPartXmlHandler(_nsm, node));
            c.Parameters = new ExcelConnectionParameters();
            _list.Add(c);
            return c;
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
                        return null;
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
