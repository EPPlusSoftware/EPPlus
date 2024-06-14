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

namespace OfficeOpenXml.RichData
{
    internal class ExcelRichValueStructureCollection
    {
        private ExcelWorkbook _wb;
        private ZipPackagePart _part;
        private Uri _uri;
        private const string PART_URI_PATH = "/xl/richData/rdrichvaluestructure.xml";
        private Dictionary<RichDataStructureFlags, int> _structures=new Dictionary<RichDataStructureFlags, int>();
        internal ExcelRichValueStructureCollection(ExcelWorkbook wb) 
        {
            _wb=wb;
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
           while(xr.Read())
            {
                if(xr.IsElementWithName("s"))
                {
                    StructureItems.Add(ReadItem(xr));
                    var structureFlag = GetStructFlag(StructureItems[StructureItems.Count - 1]);
                    if(structureFlag.HasValue)
                    {
                        _structures.Add(structureFlag.Value, StructureItems.Count - 1);
                    }
                }
                else if (xr.IsElementWithName("extLst"))
                {
                    ExtLstXml = xr.ReadInnerXml();
                }
            }

        }

        private RichDataStructureFlags? GetStructFlag(ExcelRichValueStructure rvStruct)
        {
            if (rvStruct.Keys.Any(x=>x.Name.Equals("errorType")) && rvStruct.Keys.Any(x => x.Name.Equals("subType")))
            {
                if (rvStruct.Keys.Any(x => x.Name.Equals("colOffset")) || rvStruct.Keys.Any(x => x.Name.Equals("rwOffset")))
                {
                    return RichDataStructureFlags.ErrorSpill;
                }
                else
                {
                    return RichDataStructureFlags.ErrorWithSubType;
                }
            }
            else if(rvStruct.Keys.Any(x => x.Name.Equals("propagated")))
            {
                return RichDataStructureFlags.ErrorPropagated;
            }
            return null;
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

        internal void Save(ZipOutputStream stream, CompressionLevel compressionLevel, string fileName)
        {
            stream.PutNextEntry(fileName);
            stream.CompressionLevel = (OfficeOpenXml.Packaging.Ionic.Zlib.CompressionLevel)compressionLevel;
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

        internal int GetStructureId(RichDataStructureFlags structure)
        {
            if(_structures.TryGetValue(structure, out int index))
            {
                return index;
            }
            AddStructure(structure); 
            return StructureItems.Count-1;
        }
        private void AddStructure(RichDataStructureFlags structure)
        {
            var si = new ExcelRichValueStructure();
            switch(structure)
            {
                case RichDataStructureFlags.ErrorSpill:
                    si.SetAsSpillError();
                    break;
                case RichDataStructureFlags.ErrorWithSubType:
                    si.SetAsErrorWithSubType();
                    break;
                case RichDataStructureFlags.ErrorPropagated:
                    si.SetAsPropagatedError();
                    break;
            }
            StructureItems.Add(si);
            _structures.Add(structure, StructureItems.Count - 1);
        }
        public List<ExcelRichValueStructure> StructureItems { get; }=new List<ExcelRichValueStructure>();
        public string ExtLstXml { get; set; }
    }
}
