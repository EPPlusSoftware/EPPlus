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
        public ExcelRichValueCollection(ExcelWorkbook wb, ExcelRichValueStructureCollection structures)
        {
            var r = wb.Part.GetRelationshipsByType(Relationsships.schemaRichDataValueRelationship).FirstOrDefault();
            _part = wb._package.ZipPackage.GetPart(UriHelper.ResolvePartUri(r.SourceUri, r.TargetUri));
            _wb= wb;
            _structures = structures;
            ReadXml(_part.GetStream());
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
        public List<ExcelRichValue> Items { get; }=new List<ExcelRichValue>();
        public string ExtLstXml { get; internal set; }
    }
}
