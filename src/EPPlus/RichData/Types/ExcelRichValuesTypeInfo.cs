using OfficeOpenXml.Constants;
using OfficeOpenXml.Packaging;
using OfficeOpenXml.Utils;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml;
using OfficeOpenXml.Utils.Extensions;

namespace OfficeOpenXml.RichData.Types
{
    internal class ExcelRichDataValueTypeInfo
    {
        private ExcelWorkbook _wb;
        private ZipPackagePart _part;
        public ExcelRichDataValueTypeInfo(ExcelWorkbook wb)
        {
            _wb = wb;
            var r = wb.Part.GetRelationshipsByType(Relationsships.schemaRichDataValueTypeRelationship).FirstOrDefault();
            _part = wb._package.ZipPackage.GetPart(UriHelper.ResolvePartUri(r.SourceUri, r.TargetUri));
            ReadXml(_part.GetStream());
        }

        private void ReadXml(Stream stream)
        {
            var xr = XmlReader.Create(stream);
            while(xr.Read())
            {
                if(xr.IsElementWithName("global"))
                {
                    ReadKeyFlags(xr, Global);
                }
                else if(xr.IsElementWithName("types"))
                {
                    ReadKeyFlags(xr, Global);
                }
                else if(xr.IsElementWithName("extLst"))
                {
                    ExtLstXml = xr.ReadInnerXml();
                }
                else if(xr.IsEndElementWithName("rvTypesInfo"))
                {
                    break;
                }
            }
        }
        private void ReadKeyFlags(XmlReader xr, Dictionary<string, ExcelRichTypeValueKey> values)
        {
            if(xr.ReadUntil("key"))
            {
                ReadValues(xr, values);
                return;
            }
        }

        private void ReadValues(XmlReader xr, Dictionary<string, ExcelRichTypeValueKey> values)
        {
            while(xr.IsElementWithName("key") && xr.EOF == false)
            {
                while(!xr.IsEndElementWithName("key") && xr.EOF==false)
                {
                    var item = new ExcelRichTypeValueKey(xr.GetAttribute("name"));
                    values.Add(item.Name, item);
                    while (xr.Read())
                    {
                        if(xr.IsElementWithName("flag"))
                        {
                            var flag = xr.GetAttribute("name").ToEnum<RichValueKeyFlags>();
                            if (flag.HasValue)
                            {
                                var v = xr.GetAttribute("value");
                                if (v == "1" || v.Equals("true", StringComparison.OrdinalIgnoreCase))
                                {
                                    item.Flags |= flag.Value;
                                }
                                else
                                {
                                    item.Flags &= ~flag.Value;
                                }
                            }
                        }
                        else
                        {
                            if (xr.NodeType == XmlNodeType.EndElement) xr.Read();
                            if (xr.Name != "flag") break;
                        }
                    }
                    if(xr.IsEndElementWithName("keyFlags"))
                    {
                        xr.Read(); //Move to global/types end element
                        return;
                    }
                }
            }
        }

        public Dictionary<string, ExcelRichTypeValueKey>  Global { get; set; } = new Dictionary<string, ExcelRichTypeValueKey>();
        public Dictionary<string, ExcelRichTypeValueKey> Types { get; set; } = new Dictionary<string, ExcelRichTypeValueKey>();
        public string ExtLstXml { get; set; }
    }
}
