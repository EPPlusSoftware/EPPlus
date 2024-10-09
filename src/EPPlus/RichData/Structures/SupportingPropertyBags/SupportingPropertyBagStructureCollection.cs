/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  11/11/2024         EPPlus Software AB       Initial release EPPlus 8
 *************************************************************************************************/
using OfficeOpenXml.Constants;
using OfficeOpenXml.Packaging;
using OfficeOpenXml.Packaging.Ionic.Zip;
using OfficeOpenXml.Utils;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml;

namespace OfficeOpenXml.RichData.Structures.SupportingPropertyBags
{
    internal class SupportingPropertyBagStructureCollection
    {
        private ExcelWorkbook _wb;
        private ZipPackagePart _part;
        private Uri _uri;
        private const string PART_URI_PATH = "/xl/richData/rdsupportingpropertybagstructure.xml";
        private List<SupportingPropertyBagStructure> _structures = new List<SupportingPropertyBagStructure>();

        internal SupportingPropertyBagStructureCollection(ExcelWorkbook wb)
        {
            _wb = wb;
            var r = wb.Part.GetRelationshipsByType(Relationsships.schemaRichDataSupportingPropertyBagStructureRelationship).FirstOrDefault();
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
            var xr = XmlReader.Create(stream);
            while (xr.Read())
            {
                if (xr.IsElementWithName("s"))
                {
                    _structures.Add(ReadItem(xr));
                }
            }
        }

        private SupportingPropertyBagStructure ReadItem(XmlReader xr)
        {
            var keys = new List<ExcelRichValueStructureKey>();
            while (xr.Read())
            {
                if (xr.IsElementWithName("k"))
                {
                    keys.Add(new ExcelRichValueStructureKey(xr.GetAttribute("n"), xr.GetAttribute("t")));
                }
                else if (xr.IsEndElementWithName("s"))
                {
                    break;
                }
            }
            return new SupportingPropertyBagStructure(keys);
        }

        internal void Save(ZipOutputStream stream, CompressionLevel compressionLevel, string fileName)
        {
            stream.PutNextEntry(fileName);
            stream.CompressionLevel = (Packaging.Ionic.Zlib.CompressionLevel)compressionLevel;
            var sw = new StreamWriter(stream);

            sw.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
            sw.Write($"<spbStructures xmlns=\"{Schemas.schemaRichData}\" count=\"{_structures.Count}\">");
            foreach (var item in _structures)
            {
                item.WriteXml(sw);
            }
            sw.Write("</spbStructures>");
            sw.Flush();
        }

    }
}
