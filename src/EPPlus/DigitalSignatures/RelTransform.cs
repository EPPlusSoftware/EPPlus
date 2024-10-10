using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml;

namespace OfficeOpenXml.DigitalSignatures
{
    internal class RelTransform
    {
        private readonly Type[] _inputTypes = { typeof(Stream), typeof(XmlDocument), typeof(XmlNodeList) };
        private readonly Type[] _outputTypes = { typeof(Stream) };
        private string? _xpathexpr;
        private XmlDocument? _originalDoc = new XmlDocument();
        private XmlDocument? _outputDoc;
        private XmlNamespaceManager? _nsm;
        internal string TransformXml = "<Transform Algorithm=\"http://schemas.openxmlformats.org/package/2006/RelationshipTransform\">";

        string relTransformUri = "http://schemas.openxmlformats.org/package/2006/RelationshipTransform";
        string relReference = "<mdssi:RelationshipReference xmlns:mdssi=\"http://schemas.openxmlformats.org/package/2006/digital-signature\" SourceId=\"{0}\"/>";

        internal List<string> FilterRemoveRelsWith = new List<string> { "../customXml", "docProps/", "/_xmlsignatures" };

        int RIdCount = 0;
        List<string> _idList = new List<string>();

        public RelTransform(string xml)
        {
            _originalDoc.LoadXml(xml);
            InitializeRelAlt(_originalDoc);
        }

        public RelTransform(XmlDocument doc, List<string> rIds)
        {
            _idList = rIds;
            InitializeRel(doc);
        }

        public RelTransform(string xml, List<string> rIds)
        {
            _originalDoc.LoadXml(xml);
            _idList = rIds;
            InitializeRel(_originalDoc);
        }

        void InitializeRelAlt(XmlDocument doc)
        {
            _originalDoc = doc;
            _outputDoc = new XmlDocument();
            var rootNode = _outputDoc.ImportNode(doc.DocumentElement, false);
            _outputDoc.AppendChild(rootNode);

            //var sortedIds = new List<string>(_idList);

            //sortedIds.Sort(StringComparer.Ordinal);
            //List<XmlElement> nodesToAdd = new List<XmlElement>();

            //for (int i = 0; i < _idList.Count; i++)
            //{
            //    TransformXml += string.Format(relReference, _idList[i]);
            //}

            List<XmlElement> nodesToAdd = new List<XmlElement>();

            foreach (XmlElement node in doc.DocumentElement.ChildNodes)
            {
                var targetStr = node.GetAttribute("Target") ?? "";

                if (IsValidReference(targetStr))
                {
                    var id = node.GetAttribute("Id");

                    XmlElement newNode = (XmlElement)_outputDoc.ImportNode(node, true);
                    if (!newNode.HasAttribute("TargetMode"))
                    {
                        newNode.SetAttribute("TargetMode", "Internal");
                    }
                    TransformXml += string.Format(relReference, node.GetAttribute("Id"));
                    nodesToAdd.Add(newNode);
                }
            }

            var sortedNodes = nodesToAdd.OrderBy(x => x.GetAttribute("Id"));

            foreach (XmlElement node in sortedNodes)
            {
                _outputDoc.DocumentElement.AppendChild(node);
            }

            TransformXml += "</Transform>";
        }

        bool IsValidReference(string targetStr)
        {
            foreach(string filterValue in FilterRemoveRelsWith)
            {
                if(targetStr.StartsWith(filterValue))
                {
                    return false;
                }
            }
            return true;
        }

        void InitializeRel(XmlDocument doc)
        {
            _originalDoc = doc;
            _outputDoc = new XmlDocument();
            var rootNode = _outputDoc.ImportNode(doc.DocumentElement, false);
            _outputDoc.AppendChild(rootNode);

            var sortedIds = new List<string>(_idList);

            sortedIds.Sort(StringComparer.Ordinal);
            List<XmlElement> nodesToAdd = new List<XmlElement>();

            for (int i = 0; i < _idList.Count; i++)
            {
                TransformXml += string.Format(relReference, _idList[i]);
            }

            foreach (XmlElement node in doc.DocumentElement.ChildNodes)
            {
                if (_idList != null)
                {
                    var id = node.GetAttribute("Id");
                    if (_idList.Contains(id))
                    {
                        XmlElement newNode = (XmlElement)_outputDoc.ImportNode(node, true);
                        if (!newNode.HasAttribute("TargetMode"))
                        {
                            newNode.SetAttribute("TargetMode", "Internal");
                        }
                        nodesToAdd.Add(newNode);
                    }
                }
            }

            var sortedNodes = nodesToAdd.OrderBy(x => x.GetAttribute("Id"));

            foreach (XmlElement node in sortedNodes)
            {
                _outputDoc.DocumentElement.AppendChild(node);
            }

            TransformXml += "</Transform>";
        }

        internal string GetOutputXML()
        {
            return _outputDoc.OuterXml;
        }

        internal MemoryStream GetOutputStream()
        {
            var newBytes = Encoding.Default.GetBytes(GetOutputXML());
            var transformedStream = new MemoryStream(newBytes);
            return transformedStream;
        }
    }
}

