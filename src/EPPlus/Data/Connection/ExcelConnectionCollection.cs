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
    public class ExcelConnectionCollection : IEnumerable<ExcelConnection>
    {
        ExcelPackage _package;


        List<ExcelConnection> _list = new List<ExcelConnection>();
        internal  ExcelConnectionCollection(ExcelPackage package)
        {
            _package = package;
            Part = _package.ZipPackage.GetByContentType(ContentTypes.contentTypeConnections);
            ConnectionXml = new XmlDocument();
            var nsm = _package.Workbook.NameSpaceManager;
            foreach (XmlNode node in ConnectionXml.DocumentElement.SelectNodes("d:connection", nsm)) 
            {
                _list.Add(new ExcelConnection(new ConnectionDataPartXmlHandler(nsm, node)));
            }
        }
        internal IEnumerable<ZipPackagePart> Part { get; private set; }
        public XmlDocument ConnectionXml { get; private set; }

        List<ExcelConnection> _connection = new List<ExcelConnection>();
        public IEnumerator<ExcelConnection> GetEnumerator()
        {
            return _connection.GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return _connection.GetEnumerator();
        }
        ExcelPowerQuerySettings _powerQuerySettings = null;
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
