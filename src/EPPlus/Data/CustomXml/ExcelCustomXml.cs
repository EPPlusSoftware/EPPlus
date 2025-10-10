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
    /// Represnets a custom XML part in the package.
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
            }

            var ns = new NameTable();
            var nsm = new XmlNamespaceManager(ns);
            nsm.AddNamespace("ds", "http://schemas.openxmlformats.org/officeDocument/2006/customXml");
            _xmlHelper = XmlHelperFactory.Create(nsm, PropertiesXml.DocumentElement);
            foreach (XmlElement n in _xmlHelper.GetNodes("ds:schemaRefs/ds:schemaRef"))
            {
                SchemasReferences.Add(n.Attributes["ds:uri"].Value);
            }
        }
        internal void Save()
        {
            _xmlHelper.TopNode.InnerXml = "";
            foreach (var schemaRef in SchemasReferences)
            {
                XmlElement schemaRefNode = (XmlElement)_xmlHelper.CreateNode("ds:schemaRef");
                schemaRefNode.SetAttribute("ds:uri", schemaRef);
                _xmlHelper.TopNode.AppendChild(schemaRefNode);
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