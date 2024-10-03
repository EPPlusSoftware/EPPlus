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
using OfficeOpenXml.Utils;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml;
using OfficeOpenXml.RichData.Types;
using OfficeOpenXml.Packaging.Ionic.Zip;
using OfficeOpenXml.RichData.Structures.Errors;
using OfficeOpenXml.RichData.Structures.LocalImages;

namespace OfficeOpenXml.RichData.Structures
{
    internal class ExcelRichValueStructureCollection
    {
        private ExcelWorkbook _wb;
        private ZipPackagePart _part;
        private Uri _uri;
        private const string PART_URI_PATH = "/xl/richData/rdrichvaluestructure.xml";
        private Dictionary<RichDataStructureTypes, int> _structures = new Dictionary<RichDataStructureTypes, int>();
        internal ExcelRichValueStructureCollection(ExcelWorkbook wb)
        {
            _wb = wb;
            var r = wb.Part.GetRelationshipsByType(Relationsships.schemaRichDataValueStructureRelationship).FirstOrDefault();
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
                    StructureItems.Add(ReadItem(xr));
                    var structureFlag = StructureItems[StructureItems.Count - 1].StructureType;
                    if(structureFlag != RichDataStructureTypes.Preserve)
                    {
                        _structures.Add(structureFlag, StructureItems.Count - 1);
                    }
                }
                else if (xr.IsElementWithName("extLst"))
                {
                    ExtLstXml = xr.ReadInnerXml();
                }
            }

        }

        private ExcelRichValueStructure ReadItem(XmlReader xr)
        {
            var type = xr.GetAttribute("t");
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
            return RichValueStructureFactory.Create(type, keys);
        }

        internal void Save(ZipOutputStream stream, CompressionLevel compressionLevel, string fileName)
        {
            stream.PutNextEntry(fileName);
            stream.CompressionLevel = (Packaging.Ionic.Zlib.CompressionLevel)compressionLevel;
            var sw = new StreamWriter(stream);

            sw.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
            sw.Write($"<rvStructures xmlns=\"{Schemas.schemaRichData}\" count=\"{StructureItems.Count}\">");
            foreach (var item in StructureItems)
            {
                item.WriteXml(sw);
            }
            sw.Write("</rvStructures>");
            sw.Flush();
        }

        internal void CreatePart()
        {
            if (_part == null)
            {
                _part = _wb._package.ZipPackage.CreatePart(_uri, ContentTypes.contentTypeRichDataValueStructure);
                _wb.Part.CreateRelationship(_uri, TargetMode.Internal, Relationsships.schemaRichDataValueStructureRelationship);
            }
            _part.SaveHandler = Save;
        }

        internal int GetStructureId(RichDataStructureTypes structure)
        {
            if (_structures.TryGetValue(structure, out int index))
            {
                return index;
            }
            AddStructure(structure);
            return StructureItems.Count - 1;
        }
        private void AddStructure(RichDataStructureTypes structureType)
        {
            var si = RichValueStructureFactory.Create(structureType);
            StructureItems.Add(si);
            _structures.Add(structureType, StructureItems.Count - 1);
        }
        public List<ExcelRichValueStructure> StructureItems { get; } = [];
        public string ExtLstXml { get; set; }
    }
}
