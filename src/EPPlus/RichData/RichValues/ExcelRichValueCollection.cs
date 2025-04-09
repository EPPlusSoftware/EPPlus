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
using OfficeOpenXml.RichData.RichValues.Errors;
using OfficeOpenXml.RichData.Structures;
using OfficeOpenXml.Utils.FileUtils;
using OfficeOpenXml.Utils.XML;
using System;
using System.IO;
using System.Linq;
using System.Xml;

namespace OfficeOpenXml.RichData.RichValues
{
    //MS-XLSX - 2.3.6.1
    internal partial class ExcelRichValueCollection : IndexedCollection<ExcelRichValue>
    {
        private ExcelWorkbook _wb;
        ZipPackagePart _part;
        ExcelRichValueStructureCollection _structures;
        RichDataDatabase _richDataDb;
        Uri _uri;
        internal const string PART_URI_PATH = "/xl/richData/rdrichvalue.xml";
        public ExcelRichValueCollection(ExcelWorkbook wb, RichDataDatabase richDataDb)
            : base(wb.IndexStore, RichDataEntities.RichValue)
        {
            _wb = wb;
            _richDataDb = richDataDb;
            _structures = richDataDb.Structures;
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
                    Add(ReadItem(xr));
                }
                else if (xr.IsElementWithName("extLst"))
                {
                    ExtLstXml = xr.ReadInnerXml();
                }
            }
        }

        private ExcelRichValue ReadItem(XmlReader xr)
        {
            var structureIx = int.Parse(xr.GetAttribute("s"));
            var structureId = _structures.GetIdByIndex(structureIx);
            var structure = _structures.Get(structureId);
            var item = ExcelRichValueFactory.Create(structure, structure.Id, _wb.IndexStore, _richDataDb);
            int keyIx = 0;
            while (xr.IsEndElementWithName("rv") == false)
            {
                if (xr.IsElementWithName("v"))
                {
                    if (keyIx >= structure.Keys.Count) continue;
                    var val = new ExcelRichValueValue(structure.Keys[keyIx++], xr.ReadElementContentAsString(), _wb.IndexStore);
                    _richDataDb.RichValueValues.Add(val);
                    item.Values.Add(val);
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
            item.PostProcessInitialRead();
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
            sw.Write($"<rvData xmlns=\"{Schemas.schemaRichData}\" count=\"{this.Count}\">");
            foreach (var item in this)
            {
                item.SetStructure(_richDataDb);
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

        internal void AddErrorSpill(ExcelSpillErrorValue spillError)
        {
            var item = new ErrorSpillRichValue(_wb.RichData.Db)
            {
                ColOffset = spillError.SpillColOffset,
                RwOffset = spillError.SpillRowOffset,
                SubType = 1,
                ErrorType = RichDataErrorType.Spill
            };
            Add(item);
        }

        internal void AddPropagated(eErrorType errorType)
        {
            var item = new ErrorPropagatedRichValue(_wb.RichData.Db)
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
            Add(item);
        }
        internal void AddError(eErrorType errorType, int subType)
        {
            var item = new ErrorWithSubTypeRichValue(_wb.RichData.Db)
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
            Add(item);
        }
        public string ExtLstXml { get; internal set; }

    }
}
