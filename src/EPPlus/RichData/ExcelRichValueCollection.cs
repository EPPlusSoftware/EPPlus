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
        internal ZipPackagePart Part { get { return _part; } }
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

        internal void Save(ZipOutputStream stream, CompressionLevel compressionLevel, string fileName)
        {
            stream.PutNextEntry(fileName);
            stream.CompressionLevel = (OfficeOpenXml.Packaging.Ionic.Zlib.CompressionLevel)compressionLevel;
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

        internal void CreatePart()
        {
            if (_part == null)
            {
                _uri = new Uri("/xl/richData/rdrichvalue.xml", UriKind.Relative);
                _part = _wb._package.ZipPackage.CreatePart(_uri, ContentTypes.contentTypeRichDataValue);
                _wb.Part.CreateRelationship(_uri, TargetMode.Internal, Relationsships.schemaRichDataValueRelationship);
                _part.ShouldBeSaved = false;
            }
            _part.SaveHandler = Save;
        }

        internal void AddErrorSpill(ExcelRichDataErrorValue spillError)
        {
            var structureId = _structures.GetStructureId(RichDataStructureFlags.ErrorSpill);
            var item = new ExcelRichValue(structureId);            
            item.Structure = _structures.StructureItems[item.StructureId];                        
            item.AddSpillError(spillError.SpillRowOffset, spillError.SpillColOffset, "1");
            Items.Add(item);
        }

        internal void AddPropagated(eErrorType errorType)
        {
            var structureId = _structures.GetStructureId(RichDataStructureFlags.ErrorPropagated);
            var item = new ExcelRichValue(structureId);
            item.Structure = _structures.StructureItems[item.StructureId];
            switch (errorType)
            {
                case eErrorType.Calc:
                    item.AddPropagatedError(RichDataErrorType.Calc, true);
                    break;
                case eErrorType.Spill:
                    item.AddPropagatedError(RichDataErrorType.Spill, true);
                    break;
            }
            Items.Add(item);
        }
        internal void AddError(eErrorType errorType, string subType)
        {
            var structureId = _structures.GetStructureId(RichDataStructureFlags.ErrorWithSubType);
            var item = new ExcelRichValue(structureId);
            item.Structure = _structures.StructureItems[item.StructureId];
            switch (errorType)
            {
                case eErrorType.Calc:
                    item.AddError(RichDataErrorType.Calc, subType);
                    break;
                case eErrorType.Spill:
                    item.AddError(RichDataErrorType.Spill, subType);
                    break;
                case eErrorType.Name:
                    item.AddError(RichDataErrorType.Name, subType);
                    break;

            }
            Items.Add(item);
        }

        public List<ExcelRichValue> Items { get; }=new List<ExcelRichValue>();
        public string ExtLstXml { get; internal set; }
    }
}
