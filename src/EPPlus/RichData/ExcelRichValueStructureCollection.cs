using OfficeOpenXml.Constants;
using OfficeOpenXml.Packaging;
using OfficeOpenXml.Utils;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml;

namespace OfficeOpenXml.RichData
{
    internal class ExcelRichValueStructureCollection
    {
        private ExcelWorkbook _wb;
        private ZipPackagePart _part;
        private Uri _uri;
        internal ExcelRichValueStructureCollection(ExcelWorkbook wb) 
        {
            var r = wb.Part.GetRelationshipsByType(Relationsships.schemaRichDataValueStructureRelationship).FirstOrDefault();
            if (r != null)
            {
                _uri = UriHelper.ResolvePartUri(r.SourceUri, r.TargetUri);
                if (wb._package.ZipPackage.PartExists(_uri))
                {
                    _part = wb._package.ZipPackage.GetPart(_uri);
                    ReadXml(_part.GetStream());
                }
            }
        }

        private void ReadXml(Stream stream)
        {
           var xr = XmlReader.Create(stream);
           while(xr.Read())
            {
                if(xr.IsElementWithName("s"))
                {
                    StructureItems.Add(ReadItem(xr));
                }
                else if (xr.IsElementWithName("extLst"))
                {
                    ExtLstXml = xr.ReadInnerXml();
                }
            }

        }

        private ExcelRichValueStructure ReadItem(XmlReader xr)
        {
            var item = new ExcelRichValueStructure() { Type = xr.GetAttribute("t") };
            while(xr.Read())
            {
                if(xr.IsElementWithName("k"))
                {
                    item.Keys.Add(new ExcelRichValueStructureKey(xr.GetAttribute("n"), xr.GetAttribute("t")));
                }
                else if (xr.IsEndElementWithName("s"))
                {                    
                    return item;
                }
            }
            return item;
        }

        internal void Save()
        {
            if (_part == null)
            {
                _uri = new Uri("/xl/richData/rdrichvaluestructure.xml", UriKind.Relative);
                _part = _wb._package.ZipPackage.CreatePart(_uri, ContentTypes.contentTypeRichDataValueStructure);
                _wb.Part.CreateRelationship(_uri, TargetMode.Internal, Relationsships.schemaRichDataValueStructureRelationship);
            }
            
            var stream = _part.GetStream(FileMode.Create);
            var sw = new StreamWriter(stream);
            sw.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
            sw.Write($"<rvStructures xmlns=\"{Schemas.schemaRichData}\" count=\"{StructureItems.Count}\">");
            foreach(var item in StructureItems)
            {
                item.WriteXml(sw);
            }
            sw.Write("</rvStructures>");
            sw.Flush();
        }

        public List<ExcelRichValueStructure> StructureItems { get; }=new List<ExcelRichValueStructure>();
        public string ExtLstXml { get; set; }
    }
}
