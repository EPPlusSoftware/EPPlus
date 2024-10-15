using OfficeOpenXml.Constants;
using OfficeOpenXml.Packaging;
using OfficeOpenXml.Packaging.Ionic.Zip;
using OfficeOpenXml.RichData.IndexRelations;
using OfficeOpenXml.RichData.RichValues.Errors;
using OfficeOpenXml.RichData.Structures;
using OfficeOpenXml.Utils;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml;

namespace OfficeOpenXml.RichData.RichValues
{
    //MS-XLSX - 2.3.6.1
    internal partial class ExcelRichValueCollection : IndexedCollection<ExcelRichValue>
    {
        private ExcelWorkbook _wb;
        ZipPackagePart _part;
        ExcelRichValueStructureCollection _structures;
        ExcelRichData _richData;
        Uri _uri;
        internal const string PART_URI_PATH = "/xl/richData/rdrichvalue.xml";
        public ExcelRichValueCollection(ExcelWorkbook wb, ExcelRichData richData)
            : base(richData, RichDataEntities.RichValue)
        {
            _wb = wb;
            _richData = richData;
            _structures = richData.Structures;
            var r = wb.Part.GetRelationshipsByType(Relationsships.schemaRichDataValueRelationship).FirstOrDefault();
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
            var structureId = int.Parse(xr.GetAttribute("s"));
            var structure = _structures[structureId];
            //var item = new ExcelRichValue(int.Parse(xr.GetAttribute("s")));
            var item = ExcelRichValueFactory.Create(structure, structure.Id, _richData);
            //item.Structure = _structures.StructureItems[item.StructureId];

            var keys = structure.Keys.ToNameArray();
            int keyIx = 0;
            while (xr.IsEndElementWithName("rv") == false)
            {
                if (xr.IsElementWithName("v"))
                {
                    if (keyIx >= keys.Length) continue;
                    item.SetValue(keys[keyIx++], xr.ReadElementContentAsString());
                }
                else if (xr.IsElementWithName("fb"))
                {
                    item.FallbackType = GetFBType(xr.GetAttribute("t"));
                    item.FallbackValue = xr.ReadElementContentAsString();
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
            switch (t)
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
            stream.CompressionLevel = (Packaging.Ionic.Zlib.CompressionLevel)compressionLevel;
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
                _part = _wb._package.ZipPackage.CreatePart(_uri, ContentTypes.contentTypeRichDataValue);
                _wb.Part.CreateRelationship(_uri, TargetMode.Internal, Relationsships.schemaRichDataValueRelationship);
            }
            _part.SaveHandler = Save;
        }

        //internal void UpdateStructure(ExcelRichValue rv, int structureId)
        //{
        //    rv.StructureId = structureId;
        //    rv.Structure = _structures.StructureItems[structureId];
        //}

        internal void AddErrorSpill(ExcelRichDataErrorValue spillError)
        {
            //var structureId = _structures.GetStructureId(RichDataStructureTypes.ErrorSpill);
            //var item = new ExcelRichValue(structureId);
            var item = new ErrorSpillRichValue(_wb)
            {
                ColOffset = spillError.SpillColOffset,
                RwOffset = spillError.SpillRowOffset,
                SubType = 1,
                ErrorType = RichDataErrorType.Spill
            };
            //item.Structure = _structures.StructureItems[item.StructureId];
            //item.AddSpillError(spillError.SpillRowOffset, spillError.SpillColOffset, "1");
            Items.Add(item);
        }

        internal void AddPropagated(eErrorType errorType)
        {
            //var structureId = _structures.GetStructureId(RichDataStructureTypes.ErrorPropagated);
            //var item = new ExcelRichValue(structureId);
            //item.Structure = _structures.StructureItems[item.StructureId];
            //switch (errorType)
            //{
            //    case eErrorType.Calc:
            //        item.AddPropagatedError(RichDataErrorType.Calc, true);
            //        break;
            //    case eErrorType.Spill:
            //        item.AddPropagatedError(RichDataErrorType.Spill, true);
            //        break;
            //}
            var item = new ErrorPropagatedRichValue(_wb)
            {
                Propagated = "1"
            };
            switch (errorType)
            {
                case eErrorType.Calc:
                    item.ErrorType = RichDataErrorType.Calc;
                    break;
                case eErrorType.Spill:
                    item.ErrorType = RichDataErrorType.Spill;
                    break;

            }
            Items.Add(item);
        }
        internal void AddError(eErrorType errorType, int subType)
        {
            //var structureId = _structures.GetStructureId(RichDataStructureTypes.ErrorWithSubType);
            //var item = new ExcelRichValue(structureId);
            //item.Structure = _structures.StructureItems[item.StructureId];
            //switch (errorType)
            //{
            //    case eErrorType.Calc:
            //        item.AddError(RichDataErrorType.Calc, subType);
            //        break;
            //    case eErrorType.Spill:
            //        item.AddError(RichDataErrorType.Spill, subType);
            //        break;
            //    case eErrorType.Name:
            //        item.AddError(RichDataErrorType.Name, subType);
            //        break;
            //}
            var item = new ErrorWithSubTypeRichValue(_wb)
            {
                SubType = subType
            };
            switch (errorType)
            {
                case eErrorType.Calc:
                    item.ErrorType = RichDataErrorType.Calc;
                    break;
                case eErrorType.Spill:
                    item.ErrorType = RichDataErrorType.Spill;
                    break;

            }
            Items.Add(item);
        }

        public List<ExcelRichValue> Items { get; } = new List<ExcelRichValue>();
        public string ExtLstXml { get; internal set; }

        public override RichDataEntities EntityType => RichDataEntities.RichValue;
    }
}
