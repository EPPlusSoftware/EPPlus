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
        ZipPackagePart _part;
        internal ExcelRichValueStructureCollection(ExcelWorkbook wb) 
        {
            var r = wb.Part.GetRelationshipsByType(Relationsships.schemaRichDataValueStructureRelationship).FirstOrDefault();
            _part = wb._package.ZipPackage.GetPart(UriHelper.ResolvePartUri(r.SourceUri, r.TargetUri));
            ReadXml(_part.GetStream());
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
        public List<ExcelRichValueStructure> StructureItems { get; }=new List<ExcelRichValueStructure>();
        public string ExtLstXml { get; set; }
    }
}
