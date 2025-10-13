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
using OfficeOpenXml.Drawing;
using System.Xml;

namespace OfficeOpenXml.Data.Connection
{
    internal class ConnectionDataPartXmlHandler : IDocumentPart<ExcelConnection>
    {
        private XmlHelper _xml;

        internal ConnectionDataPartXmlHandler(XmlNamespaceManager nsm, XmlNode topNode)
        {
            _xml = XmlHelperFactory.Create(nsm, topNode);
        }
        void IDocumentPart<ExcelConnection>.Read(ExcelConnection item)
        {
            item.Id = _xml.GetXmlNodeInt("@id");
            item.Name = _xml.GetXmlNodeString("@name");
            item.Description = _xml.GetXmlNodeString("@description");
            item.Credentials = _xml.GetXmlEnumNull<eCredential>("@credentials", eCredential.Integrated).Value;
            item.SourceDatabaseFile = _xml.GetXmlNodeString("@sourceFile");
            item.OdcFile = _xml.GetXmlNodeString("@odcFile");
            item.IsDeleted = _xml.GetXmlNodeBool("@deleted", false);
            item.RefreshOnLoad = _xml.GetXmlNodeBool("@refreshOnLoad", false);
            item.KeepAlive = _xml.GetXmlNodeBool("@keepAlive", false);
            item.AutomaticRefreshInterval = _xml.GetXmlNodeInt("@automaticRefreshInterval", 0);
            item.ReconnectionMethod = _xml.GetXmlEnumNull<eReconnectionMethod>("@reconnectionMethod", eReconnectionMethod.AsRequired).Value;
            item.MinimumRefreshableVersion =_xml.GetXmlNodeInt("@minimumRefreshableVersion", 0);
            item.SavePassword = _xml.GetXmlNodeBool("@savePassword", false);
            item.IsNew = _xml.GetXmlNodeBool("@new", false);
            item.IsDeleted = _xml.GetXmlNodeBool("@deleted", false);
            item.OnlyUseConnectionFile = _xml.GetXmlNodeBool("@onlyUseConnectionFile", false);
            item.IsBackground = _xml.GetXmlNodeBool("@background", false);
            item.RefreshOnLoad = _xml.GetXmlNodeBool("@refreshOnLoad", false);
            item.SaveData = _xml.GetXmlNodeBool("@saveData", false);
            item.SingleSignOnId = _xml.GetXmlNodeString("@singleSignOnId");
            item.LastRefreshVersion = _xml.GetXmlNodeInt("@refreshVersion", 0);

            if (_xml.ExistsNode("d:dbPr"))
            {
                SetDataProperties(item);

            }
            else if(_xml.ExistsNode("d:olapPr"))
            {
                SetOlapProperties(item);
            }
            else if(_xml.ExistsNode("d:webPr"))
            {
                SetWebProperties(item);
            }
            else if(_xml.ExistsNode("d:textPr"))
            {
                SetTextProperties(item);
            }
            SetParameters(item);
        }

        private void SetParameters(ExcelConnection item)
        {
            var parameters = new ExcelConnectionParameters();
            foreach(XmlElement pn in _xml.GetNodes("d:parameters/d:parameter"))
            {
                var paramXml = XmlHelperFactory.Create(_xml.NameSpaceManager, pn);
                var p = new ExcelConnectionParameter();
                p.Name = paramXml.GetXmlNodeString("@name");
                p.SqlType = paramXml.GetXmlNodeInt("@sqlType", 0);
                p.ParameterType = paramXml.GetXmlEnum("@parameterType", eConnectionParameterType.Prompt);
                p.Prompt = paramXml.GetXmlNodeString("@prompt");
                p.RefreshOnChange = paramXml.GetXmlNodeBool("@refreshOnChange", false);
                p.Boolean = paramXml.GetXmlNodeBoolNullable("@boolean");
                p.Double = paramXml.GetXmlNodeDoubleNull("@double");
                p.Integer = paramXml.GetXmlNodeIntNull("@integer");
                p.String = paramXml.GetXmlNodeString("@string", null);
                p.Cell = paramXml.GetXmlNodeString("@cell", null);
            }
            item.Parameters = parameters;
        }

        private void SetTextProperties(ExcelConnection item)
        {
            var textPr = new ExcelTextProperties();
            textPr.Prompt = _xml.GetXmlNodeBool("d:textPr/@prompt", true);
            textPr.FileType = _xml.GetXmlEnum("d:textPr/@fileType", eConnectionTextFileType.Win);
            textPr.CharacterSet = _xml.GetXmlNodeString("d:textPr/@characterSet");
            textPr.FirstRow = (uint)_xml.GetXmlNodeInt("d:textPr/@firstRow", 1);
            textPr.SourceFile = _xml.GetXmlNodeString("d:textPr/@sourceFile");
            textPr.Delimited = _xml.GetXmlNodeBool("d:textPr/@delimited", true);    
            textPr.Decimal = _xml.GetXmlNodeString("d:textPr/@decimal",".");
            textPr.Thousands = _xml.GetXmlNodeString("d:textPr/@thousands",",");
            textPr.Tab = _xml.GetXmlNodeBool("d:textPr/@tab", true);
            textPr.Space = _xml.GetXmlNodeBool("d:textPr/@space", false);
            textPr.Semicolon = _xml.GetXmlNodeBool("d:textPr/@semicolon", false);   
            textPr.Comma = _xml.GetXmlNodeBool("d:textPr/@comma", false);
            textPr.Consecutive = _xml.GetXmlNodeBool("d:textPr/@consecutive", false);
            textPr.Qualifier = _xml.GetXmlEnum("d:textPr/@qualifier", eConnectionTextQualifier.DoubleQuote);   
            textPr.Delimiter = _xml.GetXmlNodeString("d:textPr/@delimiter");
            foreach (XmlElement tfn in _xml.GetNodes("d:textPr/d:textFields"))
            {
                var tf = XmlHelperFactory.Create(_xml.NameSpaceManager, tfn);
                textPr.Fields.Add(new ExcelConnectionTextField() { Position = tf.GetXmlNodeInt("@position"), Type = tf.GetXmlNodeString("@type").TranslateExternalConnectionType() });
            }
            item.TextProperties = textPr;
        }

        private void SetWebProperties(ExcelConnection item)
        {
            var webPr = new ExcelWebProperties();
            webPr.IsXml = _xml.GetXmlNodeBool("d:webPr/@xml", false);
            webPr.IsXmlSourceData = _xml.GetXmlNodeBool("d:webPr/@sourceData");
            webPr.ParsePRE = _xml.GetXmlNodeBool("d:webPr/@parsePre", false);
            webPr.Consecutive = _xml.GetXmlNodeBool("d:webPr/@consecutive", false);
            webPr.FirstRow = _xml.GetXmlNodeBool("d:webPr/@firstRow", false);
            webPr.IsExcel97 = _xml.GetXmlNodeBool("d:webPr/@xl97", false);
            webPr.IsExcel2000 = _xml.GetXmlNodeBool("d:webPr/@xl2000", false);
            webPr.TextDates = _xml.GetXmlNodeBool("d:webPr/@textDates", false);
            webPr.Url = _xml.GetXmlNodeString("d:webPr/@url");
            webPr.Post = _xml.GetXmlNodeString("d:webPr/@post");
            webPr.HtmlTables = _xml.GetXmlNodeBool("d:webPr/@htmlTables", false);
            webPr.HtmlFormat = _xml.GetXmlEnum("d:webPr/@htmlFormat", eHtmlFormatingHandling.None);
            webPr.EditPage = _xml.GetXmlNodeString("d:webPr/editPage");

            foreach (XmlElement tn in _xml.GetNodes("d:webPr/d:tables"))
            {
                var t = XmlHelperFactory.Create(_xml.NameSpaceManager, tn);
                var ix = t.GetXmlNodeIntNull("d:ix/@v");
                if (ix.HasValue)
                {
                    webPr.Tables.Add(new ExcelHtmlTableReference() { Index = ix.Value });
                }
                else
                {
                    var s = t.GetXmlNodeString("d:s/@v");
                    if (string.IsNullOrEmpty(s) && t.ExistsNode("d:m"))
                    {
                        webPr.Tables.Add(new ExcelHtmlTableReference() { Index = -1 }); //Missiong table
                    }
                    else
                    {
                        webPr.Tables.Add(new ExcelHtmlTableReference() { Name = s });
                    }
                }
            }
            item.WebProperties = webPr;
        }

        private void SetOlapProperties(ExcelConnection item)
        {
            var olapPr = new ExcelConnectionOlapProperties();
            olapPr.Local = _xml.GetXmlNodeBool("d:olapPr/@local", false);
            olapPr.LocalConnection = _xml.GetXmlNodeString("d:olapPr/@localConnection");
            olapPr.LocalRefresh = _xml.GetXmlNodeBool("d:olapPr/@localRefresh", true);
            olapPr.SendLocale = _xml.GetXmlNodeBool("d:olapPr/@sendLocale", false);
            olapPr.ServerFill = _xml.GetXmlNodeBool("d:olapPr/@serverFill", true);
            olapPr.ServerNumberFormat = _xml.GetXmlNodeBool("d:olapPr/@serverNumberFormat", true);
            olapPr.ServerFont = _xml.GetXmlNodeBool("d:olapPr/@serverFont", true);
            olapPr.ServerFontColor = _xml.GetXmlNodeBool("d:olapPr/@serverFontColor", true);
            item.OlapProperties = olapPr;
        }

        private void SetDataProperties(ExcelConnection item)
        {
            var dbPr = new ExcelDatabaseProperties();
            dbPr.Connection = _xml.GetXmlNodeString("d:dbPr/@connection");
            dbPr.CommandType = (eCommandType)_xml.GetXmlNodeInt("d:dbPr/@commandType", 2);
            dbPr.Command = _xml.GetXmlNodeString("d:dbPr/@command");
            dbPr.ServerCommand = _xml.GetXmlNodeString("d:dbPr/@command");
            item.DatabaseProperties = dbPr;
        }

        public int Id { get; set; }
        public string Name { get; set; }
        public string Description { get; set; }
        public eCredential Credentials { get; set; } = eCredential.Integrated;
        public bool IsDeleted { get; set; }
        public bool IsBackground { get; set; }
        public int AutomaticRefreshInterval { get; set; } = 0;
        public bool KeepAlive { get; set; } = false;
        public int MinimumRefreshableVersion { get; set; } = 0;
        public bool IsNew { get; set; } = false;
        public string OdcFile { get; set; }
        public bool OnlyUseConnectionFile { get; set; } = false;
        public int LastRefreshVersion { get; set; }
        public eReconnectionMethod ReconnectionMethod { get; set; } = eReconnectionMethod.AsRequired;
        public bool RefreshOnLoad { get; set; } = false;
        public bool SaveData { get; set; } = false;
        public bool SavePassword { get; set; } = false;
        public string SingleSignOnId { get; set; }
        public string SourceDatabaseFile { get; set; }
        public eConnectionDataSourceType? Type { get; } = null;
        public ExcelDatabaseProperties DatabaseProperties { get; }
        public ExcelConnectionOlapProperties OlapProperties { get; }
        public ExcelWebProperties WebProperties { get; }
        public ExcelTextProperties TextProperties { get; }
        public ExcelConnectionParameters Parameters { get; }

        void IDocumentPart<ExcelConnection>.Save(ExcelConnection item)
        {
            throw new System.NotImplementedException();
        }
    }
}