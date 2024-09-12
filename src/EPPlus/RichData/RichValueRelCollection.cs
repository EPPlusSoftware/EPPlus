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
    internal class RichValueRelCollection
    {
        const string PART_URI_PATH = "/xl/richData/richValueRel.xml.rels";
        Uri _uri;
        private ExcelWorkbook _wb;
        ZipPackagePart _part;

        public List<RichValueRel> Items { get; } = new List<RichValueRel>();

        public RichValueRelCollection(ExcelWorkbook wb)
        {
            _wb = wb;
            var r = wb.Part.GetRelationshipsByType(Relationsships.schemaRichDataRelRelationship).FirstOrDefault();
            if (r == null)
            {
                _uri = new Uri(PART_URI_PATH, UriKind.Relative);
            }
            else
            {
                _uri = UriHelper.ResolvePartUri(r.SourceUri, r.TargetUri);
            }
            LoadPart(wb);
        }

        private void LoadPart(ExcelWorkbook wb)
        {
            if (wb._package.ZipPackage.PartExists(_uri))
            {
                _part = wb._package.ZipPackage.GetPart(_uri);
                ReadXml(_part.GetStream());
            }
        }

        internal ZipPackagePart Part { get { return _part; } }

        private void ReadXml(Stream stream)
        {
            //var ns = "http://schemas.openxmlformats.org/package/2006/relationships";
            var xml = string.Empty;
            //using var sr = new StreamReader(stream);
            //var s = sr.ReadToEnd();
            var xr = XmlReader.Create(stream);
            while (xr.Read())
            {
                if (xr.IsElementWithName("rel"))
                {
                    Items.Add(ReadItem(xr));
                }
            }
        }

        private RichValueRel ReadItem(XmlReader xr)
        {
            var ns = "http://schemas.openxmlformats.org/package/2006/relationships";
            var item = new RichValueRel();

            var id = xr.GetAttribute("r:id");
            var rel = _part.GetRelationship(id);

            item.Id = id;
            item.Type = rel.RelationshipType;
            item.Target = UriHelper.ResolvePartUri(rel.SourceUri, rel.TargetUri).OriginalString;

            return item;

        }
    }
}
