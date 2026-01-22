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
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml;

namespace OfficeOpenXml.Data.CustomXml
{
    /// <summary>
    /// Represents a custom XML part in the package.
    /// </summary>
    public class ExcelCustomXml
    {
        XmlHelper _xmlHelper;
        /// <summary>
        /// A list of the schema references that applies to the <see cref="CustomXml"/> document.
        /// </summary>
        public List<string> SchemasReferences { get; set; } = new List<string>();
        /// <summary>
        /// The custom Xml document.
        /// </summary>
        public XmlDocument CustomXml { get; set; }
        internal ZipPackagePart Part { get; set; }
        internal ZipPackagePart PropertiesPart { get; set; }
        internal XmlDocument PropertiesXml
        {
            get;
            private set;
        }
        internal ExcelCustomXml()
        {
        }
        internal ExcelCustomXml(ZipPackagePart part)
        {
            Part = part;
            var ms = part.GetStream();
            CustomXml = new XmlDocument();
            XmlHelper.LoadXmlSafe(CustomXml, ms);

            var rels = Part.GetRelationships();
            if (rels.Count > 0)
            {
                var rel = rels.First();
                PropertiesPart = part.Package.GetPart(UriHelper.ResolvePartUri(rel.SourceUri, rel.TargetUri));
                PropertiesXml = new XmlDocument();
                XmlHelper.LoadXmlSafe(PropertiesXml, PropertiesPart.GetStream());
            }
            else
            {
                PropertiesXml = null;
                return;
            }

            var nsm = CreateNsm();
            var topNode = PropertiesXml.DocumentElement.SelectSingleNode("ds:schemaRefs", nsm);
            if (topNode != null)
            {
                _xmlHelper = XmlHelperFactory.Create(nsm, topNode);
                foreach (XmlElement n in _xmlHelper.GetNodes("ds:schemaRef"))
                {
                    SchemasReferences.Add(n.Attributes["ds:uri"].Value);
                }
            }
        }

        private static XmlNamespaceManager CreateNsm()
        {
            var ns = new NameTable();
            var nsm = new XmlNamespaceManager(ns);
            nsm.AddNamespace("ds", "http://schemas.openxmlformats.org/officeDocument/2006/customXml");
            return nsm;
        }

        internal void Save(ExcelPackage pck)
        {
            if(_xmlHelper==null && SchemasReferences.Count > 0)
            {
                if(Part==null)
                {
                    var zp = pck.ZipPackage;
                    int id=1;
                    Part = zp.CreatePart(XmlHelper.GetNewUri(zp, "/customXml/item{0}.xml", ref id), string.Empty, CompressionLevel.Default, "xml");
                    PropertiesPart = zp.CreatePart(XmlHelper.GetNewUri(zp, "/customXml/itemProps{0}.xml", ref id), ContentTypes.contentTypeCustomXmlProperties);
                    Part.CreateRelationship(UriHelper.ResolvePartUri(Part.Uri, PropertiesPart.Uri), TargetMode.Internal,  $"{ExcelPackage.schemaRelationships}/customXmlProps");  
                    pck.Workbook.Part.CreateRelationship(UriHelper.ResolvePartUri(pck.Workbook.Part.Uri, Part.Uri), TargetMode.Internal, $"{ExcelPackage.schemaRelationships}/customXml");
                }
                if (PropertiesXml==null)
                {
                    PropertiesXml = new XmlDocument();
                    PropertiesXml.LoadXml( $"<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"no\"?>\r\n<ds:datastoreItem ds:itemID=\"{{{Guid.NewGuid().ToString().ToUpperInvariant()}}}\" xmlns:ds=\"http://schemas.openxmlformats.org/officeDocument/2006/customXml\"><ds:schemaRefs/></ds:datastoreItem>");
                }
                var nsm = CreateNsm();
                _xmlHelper = XmlHelperFactory.Create(nsm, PropertiesXml.DocumentElement.SelectSingleNode("ds:schemaRefs", nsm));
            }
            if (_xmlHelper != null)
            {
                _xmlHelper.TopNode.InnerXml = "";
                foreach (var schemaRef in SchemasReferences)
                {
                    XmlElement schemaRefNode = (XmlElement)_xmlHelper.CreateNode("ds:schemaRef");
                    schemaRefNode.SetAttribute("uri", CreateNsm().LookupNamespace("ds"), schemaRef);
                    _xmlHelper.TopNode.AppendChild(schemaRefNode);
                }
            }
            var xmlSettings = new XmlWriterSettings();

            var stream = Part.GetStream(FileMode.Create, FileAccess.Write);
            var xmlWriter = XmlWriter.Create(stream, xmlSettings);
            CustomXml.Save(xmlWriter);

            stream = PropertiesPart.GetStream(FileMode.Create, FileAccess.Write);
            xmlWriter = XmlWriter.Create(stream, xmlSettings);
            PropertiesXml.Save(xmlWriter);
        }
    }
}