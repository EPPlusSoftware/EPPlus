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
    //MS-XLSX - 2.3.6.1
    internal partial class ExcelRichValueCollection
    {
        private ExcelWorkbook _wb;
        ZipPackagePart _part;
        ExcelRichValueStructureCollection _structures;
        Uri _uri;
        public ExcelRichValueCollection(ExcelWorkbook wb, ExcelRichValueStructureCollection structures)
        {
            _wb = wb;
            _structures = structures;
            var r = wb.Part.GetRelationshipsByType(Relationsships.schemaRichDataValueRelationship).FirstOrDefault();
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
            while (xr.Read())
            {
                if (xr.IsElementWithName("rv"))
                {
                    Items.Add(ReadItem(xr));
                }
                else if (xr.IsElementWithName("extLst"))
                {
                    ExtLstXml = xr.ReadInnerXml();
                }
            }
        }

        private ExcelRichValue ReadItem(XmlReader xr)
        {
            var item = new ExcelRichValue(int.Parse(xr.GetAttribute("s")));
            item.Structure = _structures.StructureItems[item.StructureId];

            while (xr.IsEndElementWithName("rv")==false)
            {
                if (xr.IsElementWithName("v"))
                {
                    item.Values.Add(xr.ReadElementContentAsString());
                }
                else if (xr.IsElementWithName("fb"))
                {
                    item.Fallback = GetFBType(xr.GetAttribute("t"));
                    xr.Read();
                }
                else 
                {
                    xr.Read();
                }

            }
            return item;
        }
        private RichValueFallbackType GetFBType(string t)
        {
            switch(t)
            {
                case "b":
                    return RichValueFallbackType.Boolean;
                case "e":
                    return RichValueFallbackType.Error;
                case "s":
                    return RichValueFallbackType.String;
                default:
                    return RichValueFallbackType.Decimal;
            }
        }

        internal void Save()
        {
            if (_part == null)
            {
                _uri = new Uri("/xl/richData/rdrichvalue.xml", UriKind.Relative);
                _part = _wb._package.ZipPackage.CreatePart(_uri, ContentTypes.contentTypeRichDataValue);
                _wb.Part.CreateRelationship(_uri, TargetMode.Internal, Relationsships.schemaRichDataValueRelationship);
            }

            var stream = _part.GetStream(FileMode.Create);
            var sw = new StreamWriter(stream);
            sw.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
            sw.Write($"<rvData xmlns=\"{Schemas.schemaRichData}\" count=\"{Items.Count}\">");
            foreach (var item in Items)
            {
                item.WriteXml(sw);
            }
            sw.Write("</rvData>");
            sw.Flush();
        }

        public List<ExcelRichValue> Items { get; }=new List<ExcelRichValue>();
        public string ExtLstXml { get; internal set; }
    }
}
