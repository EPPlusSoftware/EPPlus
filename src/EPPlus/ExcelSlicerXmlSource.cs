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
using OfficeOpenXml.Packaging;
using OfficeOpenXml.Utils;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml;

namespace OfficeOpenXml
{
    internal class ExcelSlicerXmlSources : XmlHelper
    {
        const string _tableUId = "{3A4CF648-6AED-40f4-86FF-DC5316D8AED3}";
        const string _pivotTableUId = "{A8765BA9-456A-4dab-B4F3-ACF838C121DE}";
        internal List<ExcelSlicerXmlSource> _list = new List<ExcelSlicerXmlSource>();
        internal ExcelSlicerXmlSources(XmlNamespaceManager nsm, XmlNode topNode, ZipPackagePart part) : base(nsm, topNode)
        {
            foreach (XmlNode node in GetNodes("d:extLst/d:ext"))
            {
                switch (node.Attributes["uri"].Value)
                {
                    case _tableUId:  //Table slicer
                        foreach (XmlNode slicerNode in node.SelectNodes("x14:slicerList/x14:slicer", NameSpaceManager))
                        {
                            _list.Add(new ExcelSlicerXmlSource(eSlicerSourceType.Table, part, slicerNode.Attributes["r:id"].Value));
                        }
                        break;
                    case _pivotTableUId: //Pivot table slicer
                        foreach (XmlNode slicerNode in node.SelectNodes("x14:slicerList/x14:slicer", NameSpaceManager))
                        {
                            _list.Add(new ExcelSlicerXmlSource(eSlicerSourceType.PivotTable, part, slicerNode.Attributes["r:id"].Value));
                        }
                        break;
                    default:
                        break;
                }
            }
        }
        internal ExcelSlicerXmlSource GetOrCreateSource(eSlicerSourceType sourceType)
        {
            var src = GetSources(sourceType).FirstOrDefault();
            if(src==null)
            {
                switch(sourceType)
                {
                    case eSlicerSourceType.Table:
                        break;
                    case eSlicerSourceType.PivotTable:
                        break;
                }
            }
            return src;
        }
        internal XmlNode GetSource(string name, eSlicerSourceType sourceType)
        {
            foreach (var source in GetSources(sourceType))
            {
                var n = source.XmlDocument.DocumentElement.SelectSingleNode($"x14:slicer[@name=\"{name}\"]", NameSpaceManager);
                if (n != null)
                {
                    return n;
                }
            }
            return null;
        }

        private IEnumerable<ExcelSlicerXmlSource> GetSources(eSlicerSourceType sourceType)
        {
            return _list.Where(x => x.Type == sourceType);
        }
    }

    internal class ExcelSlicerXmlSource : ExcelXmlSource
    {
        internal ExcelSlicerXmlSource(eSlicerSourceType type, ZipPackagePart relPart, string relId) : base(relPart, relId)
        {
            Type = type;
        }
        public eSlicerSourceType Type { get; }
    }
    internal class ExcelXmlSource
    {
        internal ExcelXmlSource(ZipPackagePart relPart, string relId)  
        {
            Rel = relPart.GetRelationship(relId);
            Uri = UriHelper.ResolvePartUri(relPart.Uri, Rel.TargetUri);
            Part = relPart.Package.GetPart(Uri);

            var xml = new XmlDocument();
            XmlHelper.LoadXmlSafe(xml, Part.GetStream());
            XmlDocument = xml;
        }
        internal ZipPackageRelationship Rel
        {
            get;
        }
        internal ZipPackagePart Part
        {
            get;
        }
        internal Uri Uri
        {
            get;
        }
        public XmlDocument XmlDocument
        {
            get;
        }

    }
}