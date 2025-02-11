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
using OfficeOpenXml.RichData.IndexRelations;
using OfficeOpenXml.RichData.Structures.Constants;
using OfficeOpenXml.Utils;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml;

namespace OfficeOpenXml.RichData.Structures
{
    internal class ExcelRichValueStructureCollection : IndexedCollection<ExcelRichValueStructure>
    {
        private ExcelWorkbook _wb;
        private ZipPackagePart _part;
        private RichDataDatabase _richDataDb;
        private readonly StructureKeyNamesCache _keyNamesCache = new StructureKeyNamesCache();
        private Uri _uri;
        private const string PART_URI_PATH = "/xl/richData/rdrichvaluestructure.xml";
        private Dictionary<RichDataStructureTypes, List<RichValueStructureReference>> _structures = new Dictionary<RichDataStructureTypes, List<RichValueStructureReference>>();
        internal ExcelRichValueStructureCollection(ExcelWorkbook wb, RichDataDatabase richDataDb)
            : base(wb.IndexStore, RichDataEntities.RichStructure)
        {
            _wb = wb;
            _richDataDb = richDataDb;
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

        private void AddStructure(RichDataStructureTypes structureType, uint id, List<int> ids)
        {
            if(!_structures.ContainsKey(structureType))
            {
                _structures[structureType] = new List<RichValueStructureReference>();
            }
            var reference = new RichValueStructureReference(id);
            reference.WordIds.AddRange(ids);
            _structures[structureType].Add(reference);
        }

        internal ZipPackagePart Part { get { return _part; } }
        private void ReadXml(Stream stream)
        {
            var xr = XmlReader.Create(stream);
            while (xr.Read())
            {
                if (xr.IsElementWithName("s"))
                {
                    var structure = ReadItem(xr);
                    Add(structure);
                    var structureFlag = this[Count - 1].StructureType;
                    if(structureFlag != RichDataStructureTypes.Preserve)
                    {
                        var ids = _keyNamesCache.GetIds(structure.Keys.Select(k => k.Name));
                        AddStructure(structureFlag, structure.Id, ids);
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
            return RichValueStructureFactory.Create(type, keys, _wb.IndexStore);
        }

        internal void Save(ZipOutputStream stream, CompressionLevel compressionLevel, string fileName)
        {
            stream.PutNextEntry(fileName);
            stream.CompressionLevel = (Packaging.Ionic.Zlib.CompressionLevel)compressionLevel;
            var sw = new StreamWriter(stream);

            sw.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
            sw.Write($"<rvStructures xmlns=\"{Schemas.schemaRichData}\" count=\"{Count}\">");
            foreach (var item in this)
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

        private uint AddStructure(RichDataStructureTypes structureType)
        {
            var si = RichValueStructureFactory.Create(structureType, _wb.IndexStore);
            Add(si);
            var ids = _keyNamesCache.GetIds(si.Keys.Select(k => k.Name));
            AddStructure(structureType, si.Id, ids);
            return si.Id;
        }

        internal ExcelRichValueStructure GetByType(RichDataStructureTypes structure, List<string> keyNames)
        {
            var keyIds = _keyNamesCache.GetIds(keyNames);
            var sk = structure;
            if((sk & RichDataStructureTypes.Error) == RichDataStructureTypes.Error)
            {
                sk = RichDataStructureTypes.Error;
            }
            if (_structures.TryGetValue(sk, out List<RichValueStructureReference> structureRefs))
            {
                foreach(var reference in structureRefs)
                {
                    if(reference.AreEqual(keyIds))
                    {
                        return Get(reference.Id);
                    }
                }
            }
            var keys = new List<ExcelRichValueStructureKey>();
            var structureName = StructureTypes.GetStructureName(structure);
            StructureKeys.SortKeyNames(structure, ref keyNames);
            foreach (var key in keyNames)
            {
                var dt = StructureKeys.GetKeyDataType(structureName, key);
                if(dt.HasValue)
                {
                    var structureKey = new ExcelRichValueStructureKey(key, dt.Value);
                    keys.Add(structureKey);
                }
            }
            var rvStructure = RichValueStructureFactory.Create(structure, keys, _wb.IndexStore);
            var keyNames2 = rvStructure.Keys.Select(k => k.Name);
            if(!_structures.ContainsKey(sk))
            {
                _structures[sk] = new List<RichValueStructureReference>();
            }
            var newStructureRef = new RichValueStructureReference(rvStructure.Id);
            newStructureRef.WordIds.AddRange(_keyNamesCache.GetIds(keyNames2));
            _structures[sk].Add(newStructureRef);
            Add(rvStructure);
            return rvStructure;
        }

        public string ExtLstXml { get; set; }
    }
}
