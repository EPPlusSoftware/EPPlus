﻿/*************************************************************************************************
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
using OfficeOpenXml.RichData.IndexRelations;

namespace OfficeOpenXml.RichData.Structures
{
    internal class ExcelRichValueStructureCollection : IndexedCollection<ExcelRichValueStructure>
    {
        private ExcelWorkbook _wb;
        private ZipPackagePart _part;
        private ExcelRichData _richData;
        private Uri _uri;
        private const string PART_URI_PATH = "/xl/richData/rdrichvaluestructure.xml";
        private Dictionary<RichDataStructureTypes, int> _structures = new Dictionary<RichDataStructureTypes, int>();
        internal ExcelRichValueStructureCollection(ExcelWorkbook wb, ExcelRichData richData)
            : base(wb.IndexStore, RichDataEntities.RichStructure)
        {
            _wb = wb;
            _richData = richData;
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

        public override RichDataEntities EntityType => RichDataEntities.RichStructure;

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
                    //StructureItems.Add(ReadItem(xr));
                    var structure = ReadItem(xr);
                    Add(structure);
                    //var structureFlag = StructureItems[StructureItems.Count - 1].StructureType;
                    var structureFlag = this[Count - 1].StructureType;
                    if(structureFlag != RichDataStructureTypes.Preserve)
                    {
                        //_structures.Add(structureFlag, StructureItems.Count - 1);
                        _structures.Add(structureFlag, structure.Id);
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
            return RichValueStructureFactory.Create(type, keys, _richData);
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

        internal int GetStructureId(RichDataStructureTypes structure)
        {
            if (_structures.TryGetValue(structure, out int index))
            {
                //index = structure.Id;
                return index;
            }
            return AddStructure(structure);
            //return StructureItems.Count - 1;
        }

        internal ExcelRichValueStructure GetByType(RichDataStructureTypes structure)
        {
            if (_structures.TryGetValue(structure, out int id))
            {
                return GetItemById(id);
            }
            var id2 = AddStructure(structure);
            return GetItemById(id2);
            //return default;
        }
        private int AddStructure(RichDataStructureTypes structureType)
        {
            var si = RichValueStructureFactory.Create(structureType, _wb.RichData);
            //StructureItems.Add(si);
            Add(si);
            //_structures.Add(structureType, StructureItems.Count - 1);
            _structures.Add(structureType, si.Id);
            return si.Id;
        }
        //public List<ExcelRichValueStructure> StructureItems { get; } = [];
        public string ExtLstXml { get; set; }
    }
}
