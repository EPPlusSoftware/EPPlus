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
using OfficeOpenXml.Core;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Utils.Compare;
using OfficeOpenXml.Utils.EnumUtils;
using System;
using System.Globalization;
using System.Net;
using System.Xml;

namespace OfficeOpenXml.Data.Connection
{
    internal class ConnectionDataPartXmlHandler : IDocumentPart<ExcelConnection>
    {
        private XmlHelper _xml;
        internal ConnectionDataPartXmlHandler(XmlNamespaceManager nsm, XmlNode topNode)
        {
            _xml = XmlHelperFactory.Create(nsm, topNode);
            _xml.SchemaNodeOrder = ["dbPr", "olapPr", "webPr", "textPr", "parameters"];
        }
        void IDocumentPart<ExcelConnection>.Load(ExcelConnection item)
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
                LoadDataProperties(item);

            }
            else if(_xml.ExistsNode("d:olapPr"))
            {
                LoadOlapProperties(item);
            }
            else if(_xml.ExistsNode("d:webPr"))
            {
                LoadWebProperties(item);
            }
            else if(_xml.ExistsNode("d:textPr"))
            {
                LoadTextProperties(item);
            }
            LoadParameters(item);            
        }
        void IDocumentPart<ExcelConnection>.Remove()
        {
            _xml.TopNode.ParentNode.RemoveChild(_xml.TopNode);
        }

        private void LoadParameters(ExcelConnection item)
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

        private void LoadTextProperties(ExcelConnection item)
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
                textPr.Fields.Add(new ExcelConnectionTextField(tf.GetXmlNodeString("@type").TranslateExternalConnectionType(), tf.GetXmlNodeInt("@position")));
            }
            item.TextProperties = textPr;
        }

        private void LoadWebProperties(ExcelConnection item)
        {
            var webPr = new ExcelWebProperties();
            webPr.IsXml = _xml.GetXmlNodeBool("d:webPr/@xml", false);
            webPr.ParsePRE = _xml.GetXmlNodeBool("d:webPr/@parsePre", false);
            webPr.IsXmlSourceData = _xml.GetXmlNodeBool("d:webPr/@sourceData");
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
                        webPr.Tables.Add(new ExcelHtmlTableReference() { Index = -1 }); //Missing table
                    }
                    else
                    {
                        webPr.Tables.Add(new ExcelHtmlTableReference() { Name = s });
                    }
                }
            }
            item.WebProperties = webPr;
        }

        private void LoadOlapProperties(ExcelConnection item)
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

        private void LoadDataProperties(ExcelConnection item)
        {
            var dbPr = new ExcelDatabaseProperties();
            dbPr.Connection = _xml.GetXmlNodeString("d:dbPr/@connection");
            dbPr.CommandType = (eCommandType)_xml.GetXmlNodeInt("d:dbPr/@commandType", 2);
            dbPr.Command = _xml.GetXmlNodeString("d:dbPr/@command");
            dbPr.ServerCommand = _xml.GetXmlNodeString("d:dbPr/@serverCommand", null);
            item.DatabaseProperties = dbPr;
        }
        void IDocumentPart<ExcelConnection>.Save(ExcelConnection item)
        {
            _xml.SetXmlNodeInt("@id", item.Id);
            _xml.SetXmlNodeString("@name", item.Name);
            _xml.SetXmlNodeString("@description", item.Description, true);
            _xml.SetXmlNodeString("@credentials", item.Credentials.ToEnumString(eCredential.Integrated), true);
            _xml.SetXmlNodeString("@sourceFile", item.SourceDatabaseFile, true);
            _xml.SetXmlNodeString("@odcFile", item.OdcFile, true);
            _xml.SetXmlNodeBool("@deleted", item.IsDeleted, false);
            _xml.SetXmlNodeBool("@refreshOnLoad",item.RefreshOnLoad, false);
            _xml.SetXmlNodeBool("@keepAlive", item.KeepAlive, false);
            _xml.SetXmlNodeInt("@automaticRefreshInterval", item.AutomaticRefreshInterval, 0);
            _xml.SetXmlNodeString("@reconnectionMethod", item.ReconnectionMethod.ToEnumString(eReconnectionMethod.AsRequired), true);
            _xml.SetXmlNodeInt("@minimumRefreshableVersion", item.MinimumRefreshableVersion, 0);
            _xml.SetXmlNodeBool("@savePassword", item.SavePassword, false);
            _xml.SetXmlNodeBool("@new", item.IsNew, false);
            _xml.SetXmlNodeBool("@deleted", item.IsDeleted, false);
            _xml.SetXmlNodeBool("@onlyUseConnectionFile", item.OnlyUseConnectionFile, false);
            _xml.SetXmlNodeBool("@background", item.IsBackground, false);
            _xml.SetXmlNodeBool("@refreshOnLoad", item.RefreshOnLoad, false);
            _xml.SetXmlNodeBool("@saveData", item.SaveData, false);
            _xml.SetXmlNodeString("@singleSignOnId", item.SingleSignOnId, true);
            _xml.SetXmlNodeInt("@refreshVersion", item.LastRefreshVersion, 0);

            SaveDataProperties(item.DatabaseProperties);
            SaveOlapProperties(item.OlapProperties);
            SaveWebProperties(item.WebProperties);
            SaveTextProperties(item.TextProperties);
            SaveParameters(item.Parameters);
        }

        private void SaveParameters(ExcelConnectionParameters parameters)
        {
            if (parameters.Count == 0)
            {
                _xml.DeleteNode("d:parameters");
            }
            else
            {
                var parametersNode = (XmlElement)_xml.CreateNode("d:parameters");
                parametersNode.RemoveAll();
                foreach(var p in parameters)
                {
                    var pn = _xml.TopNode.OwnerDocument.CreateElement("d:Parameter", Schemas.schemaMain);
                    parametersNode.AppendChild(pn);
                    if(string.IsNullOrEmpty(p.Name)==false) pn.SetAttribute("name", p.Name);
                    if (p.SqlType != 0) pn.SetAttribute("sqlType", p.SqlType.ToString());
                    if (p.ParameterType != eConnectionParameterType.Prompt) pn.SetAttribute("parameterType", p.ParameterType.ToEnumString());
                    if(string.IsNullOrEmpty(p.Prompt)==false) pn.SetAttribute("prompt", p.Prompt);
                    if(p.RefreshOnChange) pn.SetAttribute("@refreshOnChange", "1");
                    if (p.Boolean.HasValue) pn.SetAttribute("boolean", p.Boolean.Value ? "1" :  "0");
                    if(p.Double.HasValue) pn.SetAttribute("double", p.Double.Value.ToString(CultureInfo.InvariantCulture));
                    if(p.Integer.HasValue) pn.SetAttribute("double", p.Integer.Value.ToString(CultureInfo.InvariantCulture)); ;
                    if(string.IsNullOrEmpty(p.String)==false) pn.SetAttribute("string", p.String);
                    if (string.IsNullOrEmpty(p.Cell) == false) pn.SetAttribute("cell", p.Cell);
                }
                parametersNode.SetAttribute("count", parameters.Count.ToString(CultureInfo.InvariantCulture));
            }
        }
        private void SaveTextProperties(ExcelTextProperties textPr)
        {
            if (textPr == null) return;
            _xml.SetXmlNodeBool("d:textPr/@prompt", textPr.Prompt, true);
            _xml.SetXmlNodeString("d:textPr/@fileType", textPr.FileType.ToEnumString(eConnectionTextFileType.Win));
            _xml.SetXmlNodeString("d:textPr/@characterSet", textPr.CharacterSet, true);
            _xml.SetXmlNodeInt("d:textPr/@firstRow", (int)textPr.FirstRow, 1);
            _xml.SetXmlNodeString("d:textPr/@sourceFile", textPr.SourceFile, true);
            _xml.SetXmlNodeBool("d:textPr/@delimited", textPr.Delimited, true);
            _xml.SetXmlNodeString("d:textPr/@decimal", textPr.Decimal, true);
            _xml.SetXmlNodeString("d:textPr/@thousands", textPr.Thousands, true);
            _xml.SetXmlNodeBool("d:textPr/@tab",textPr.Tab, true);
            _xml.SetXmlNodeBool("d:textPr/@space", textPr.Space, false);
            _xml.SetXmlNodeBool("d:textPr/@semicolon", textPr.Semicolon, false);
            _xml.SetXmlNodeBool("d:textPr/@comma", textPr.Comma, false);
            _xml.SetXmlNodeBool("d:textPr/@consecutive", textPr.Consecutive, false);
            _xml.SetXmlNodeString("d:textPr/@qualifier", textPr.Qualifier.ToEnumString(eConnectionTextQualifier.DoubleQuote), true);
            _xml.SetXmlNodeString("d:textPr/@delimiter", textPr.Delimiter, true);
            if (textPr.Fields.Count == 0)
            {
                _xml.DeleteNode("d:textPr/d:textFields");
            }
            else
            {
                var fieldsNode = (XmlElement)_xml.CreateNode("d:textPr/d:textFields");
                fieldsNode.RemoveAll();
                foreach (var tf in textPr.Fields)
                {
                    var node = fieldsNode.OwnerDocument.CreateElement("d", _xml.NameSpaceManager.LookupNamespace("d"));
                    node.SetAttribute("type", tf.Type.ToEnumString());
                    node.SetAttribute("position", tf.Position.ToString(CultureInfo.InvariantCulture));
                    fieldsNode.AppendChild(node);
                }
                fieldsNode.SetAttribute("count", textPr.Fields.Count.ToString(CultureInfo.InvariantCulture));
            }
        }

        private void SaveWebProperties(ExcelWebProperties webPr)
        {
            if (webPr == null) return;
            _xml.SetXmlNodeBool("d:webPr/@xml", webPr.IsXml, false);
            _xml.SetXmlNodeBool("d:webPr/@sourceData", webPr.IsXmlSourceData, false);
            _xml.SetXmlNodeBool("d:webPr/@parsePre", webPr.ParsePRE, false);
            _xml.SetXmlNodeBool("d:webPr/@consecutive", webPr.Consecutive, false);
            _xml.SetXmlNodeBool("d:webPr/@firstRow", webPr.FirstRow, false);
            _xml.SetXmlNodeBool("d:webPr/@xl97", webPr.IsExcel97, false);
            _xml.SetXmlNodeBool("d:webPr/@xl2000", webPr.IsExcel2000, false);
            _xml.SetXmlNodeBool("d:webPr/@textDates", webPr.TextDates, false);
            _xml.SetXmlNodeString("d:webPr/@url", webPr.Url, true);
            _xml.SetXmlNodeString("d:webPr/@post", webPr.Post, true);
            _xml.SetXmlNodeBool("d:webPr/@htmlTables", webPr.HtmlTables, false);
            _xml.SetXmlNodeString("d:webPr/@htmlFormat", webPr.HtmlFormat.ToEnumString(eHtmlFormatingHandling.None), true);
            _xml.SetXmlNodeString("d:webPr/editPage", webPr.EditPage, true);

            if(webPr.Tables.Count==0)
            {
                _xml.DeleteNode("d:webPr/d:tables");
            }
            else
            {
                var tablesNode = (XmlElement)_xml.CreateNode("d:webPr/d:tables");
                tablesNode.RemoveAll();
                foreach (var t in webPr.Tables)
                {
                    XmlNode tableNode;
                    if(t.Index>=0)
                    {
                        tableNode = tablesNode.OwnerDocument.CreateElement("d:x", _xml.NameSpaceManager.LookupNamespace("d"));
                        tableNode.Attributes.Append(tablesNode.OwnerDocument.CreateAttribute("v"));
                        tableNode.Attributes[0].Value = t.Index.ToString(CultureInfo.InvariantCulture);
                    }
                    else if(string.IsNullOrEmpty(t.Name))
                    {
                        tableNode = tablesNode.OwnerDocument.CreateElement("d:s", _xml.NameSpaceManager.LookupNamespace("d"));
                        tableNode.Attributes.Append(tablesNode.OwnerDocument.CreateAttribute("v"));
                        tableNode.Attributes[0].Value = t.Name;
                    }
                    else
                    {
                        tableNode = tablesNode.OwnerDocument.CreateElement("d:m", _xml.NameSpaceManager.LookupNamespace("d"));
                    }
                    tablesNode.AppendChild(tableNode);
                }
                tablesNode.SetAttribute("count", webPr.Tables.Count.ToString(CultureInfo.InvariantCulture));

            }
        }

        private void SaveOlapProperties(ExcelConnectionOlapProperties olapPr)
        {
            if (olapPr == null) return;
            _xml.SetXmlNodeBool("d:olapPr/@local", olapPr.Local, false);
            _xml.SetXmlNodeString("d:olapPr/@localConnection", olapPr.LocalConnection, true);
            _xml.SetXmlNodeBool("d:olapPr/@localRefresh", olapPr.LocalRefresh, true);
            _xml.SetXmlNodeBool("d:olapPr/@sendLocale", olapPr.SendLocale, false);
            _xml.SetXmlNodeBool("d:olapPr/@serverFill", olapPr.ServerFill, true);
            _xml.SetXmlNodeBool("d:olapPr/@serverNumberFormat", olapPr.ServerNumberFormat, true);
            _xml.SetXmlNodeBool("d:olapPr/@serverFont", olapPr.ServerFont, true);
            _xml.SetXmlNodeBool("d:olapPr/@serverFontColor", olapPr.ServerFontColor, true);
        }

        private void SaveDataProperties(ExcelDatabaseProperties dbPr)
        {
            if (dbPr == null) return;
            if(string.IsNullOrEmpty(dbPr.Connection))
            {
                throw new InvalidOperationException("A connection string is required in order to save the connection.");
            }
            _xml.SetXmlNodeString("d:dbPr/@connection", dbPr.Connection);
            _xml.SetXmlNodeInt("d:dbPr/@commandType", (int)dbPr.CommandType, 2);
            _xml.SetXmlNodeString("d:dbPr/@command", dbPr.Command);
            _xml.SetXmlNodeString("d:dbPr/@serverCommand", dbPr.ServerCommand, true, false, true);
        }
    }
}