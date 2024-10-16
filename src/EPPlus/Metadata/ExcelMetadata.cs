/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  07/25/2024         EPPlus Software AB       EPPlus 7
 *************************************************************************************************/
using OfficeOpenXml.Constants;
using OfficeOpenXml.Metadata.FutureMetadata;
using OfficeOpenXml.Packaging;
using OfficeOpenXml.Packaging.Ionic.Zip;
using OfficeOpenXml.RichData;
using OfficeOpenXml.RichData.IndexRelations;
using OfficeOpenXml.RichData.RichValues.Errors;
using OfficeOpenXml.Utils;
using System;
using System.Collections.Generic;
using System.Data.Common;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml;
using static OfficeOpenXml.ExcelWorksheet;

namespace OfficeOpenXml.Metadata
{
    internal class ExcelMetadata
    {
        private ExcelWorkbook _wb;
        private ZipPackagePart _part;
        private Uri _uri;
        
        //Preserve xml variables
        private string _metadataStringsXml;
        private string _metadataStringCount;
        private string _mdxMetadataXml;
        private string _mdxMetadataCount;
        public string _extLstXml;

        internal const string FUTURE_METADATA_DYNAMIC_ARRAY_NAME = "XLDAPR";
        internal const string FUTURE_METADATA_RICHDATA_NAME = "XLRICHVALUE";

        internal MetadataTypesCollection MetadataTypes { get; }
        //internal Dictionary<string, ExcelFutureMetadata> FutureMetadata { get; } = new Dictionary<string, ExcelFutureMetadata>();
        internal FutureMetadataCollection FutureMetadata { get; set; }

        internal FutureMetadataRichValue FutureMetadataRichValue { get; private set; }

        private readonly HashSet<string> _metadataTypeNames = new HashSet<string>();

        internal FutureMetadataDynamicArray FutureMetadataDynamicArray { get; private set; }
        internal List<ExcelCellMetadataBlock> CellMetadata { get; } = new List<ExcelCellMetadataBlock>();
        internal ValueMetadataBlockCollection ValueMetadata { get; }
        internal int RichDataTypeIndex { get; private set; }
        internal int DynamicArrayTypeIndex { get; private set; }
        internal ZipPackagePart Part { get { return _part; } }

        public ExcelMetadata(ExcelWorkbook workbook)
        {
            _wb = workbook;
            var p = _wb._package;
            ValueMetadata = new ValueMetadataBlockCollection(_wb.RichData);
            MetadataTypes = new MetadataTypesCollection(_wb.RichData);
            var rel = _wb.Part.GetRelationshipsByType(ExcelPackage.schemaMetadata).FirstOrDefault();
            if(rel!=null)
            {
                _uri = UriHelper.ResolvePartUri(rel.SourceUri, rel.TargetUri);
                _part = p.ZipPackage.GetPart(_uri);
                ReadMetadata(_part.GetStream());
            }
            else
            {
                _uri = new Uri("/xl/metadata.xml", UriKind.Relative);
            }
        }

        private void ReadMetadata(Stream stream)
        {
            var xr = XmlReader.Create(stream);
            while(xr.Read())
            {
                if(xr.NodeType== XmlNodeType.Element)
                {
                    switch (xr.Name)
                    {
                        case "metadataTypes":
                            ReadMetadataTypes(xr);
                            break;
                        case "metadataStrings":
                            //Currently not used. Preserve.
                            _metadataStringsXml = xr.ReadInnerXml();
                            _metadataStringCount = xr.GetAttribute("count");
                            break;
                        case "mdxMetadata":
                            //Currently not used. Preserve.
                            _mdxMetadataXml = xr.ReadInnerXml();
                            _mdxMetadataCount = xr.GetAttribute("count");
                            break;
                        case "futureMetadata":
                            ReadFutureMetadata(xr);
                            break;
                        case "cellMetadata":
                            ReadCellMetadataItems(xr, xr.Name, CellMetadata);
                            break;
                        case "valueMetadata":
                            ReadValueMetadataItems(xr, xr.Name, ValueMetadata);
                            break;
                        case "extLst":
                            _extLstXml = xr.ReadInnerXml();
                            break;
                    }

                }
            }
        }

        private void ReadCellMetadataItems(XmlReader xr, string elementName, List<ExcelCellMetadataBlock> collection)
        {
            xr.Read();
            while(xr.IsEndElementWithName(elementName) ==false && xr.EOF==false)
            {
                if(xr.IsElementWithName("bk"))
                {
                    xr.Read();
                    while(xr.IsEndElementWithName("bk")==false)
                    {
                        collection.Add(new ExcelCellMetadataBlock(xr));
                    }
                }
                xr.Read();
            }
        }

        private void ReadValueMetadataItems(XmlReader xr, string elementName, ValueMetadataBlockCollection collection)
        {
            xr.Read();
            while (xr.IsEndElementWithName(elementName) == false && xr.EOF == false)
            {
                if (xr.IsElementWithName("bk"))
                {
                    xr.Read();
                    while (xr.IsEndElementWithName("bk") == false)
                    {
                        collection.Add(new ExcelValueMetadataBlock(xr, this, _wb.RichData.IndexStore));
                    }
                }
                xr.Read();
            }
        }

        private void ReadFutureMetadata(XmlReader xr)
        {
            //var item = new ExcelFutureMetadata();
            //item.Name = xr.GetAttribute("name");
            ExcelFutureMetadata fd;
            var name = xr.GetAttribute("name");
            if (name == FUTURE_METADATA_RICHDATA_NAME)
            {
                var fdrv = new FutureMetadataRichValue(name, _wb.RichData, this);
                FutureMetadataRichValue = fdrv;
                RichDataTypeIndex = MetadataTypes.Count;
                fd = fdrv;
            }
            else if (name == FUTURE_METADATA_DYNAMIC_ARRAY_NAME)
            {
                fd = new FutureMetadataDynamicArray(xr, _wb.RichData);
            }
            else
            {
                var count = int.Parse(xr.GetAttribute("count"));
                fd = new FutureMetadataPreserve(name, count, _wb.RichData);
            }
            if(!_metadataTypeNames.Contains(name))
            {
                _metadataTypeNames.Add(name);
            }
            fd.Index = FutureMetadata.Count;
            //FutureMetadata.Add(item.Name, item);
            FutureMetadata.Add(fd);

            //while (xr.IsEndElementWithName("futureMetadata") == false && xr.EOF == false)
            //{
            //    if (xr.IsElementWithName("ext"))
            //    {
            //        switch (xr.GetAttribute("uri"))
            //        {
            //            case ExtLstUris.DynamicArrayPropertiesUri:
            //                item.Types.Add(new ExcelFutureMetadataDynamicArray(xr));
            //                break;
            //            case ExtLstUris.RichValueDataUri:
            //                item.Types.Add(new ExcelFutureMetadataRichData(xr));
            //                break;
            //        }                    
            //    }
            //    xr.Read();
            //}
        }
        private void ReadMetadataTypes(XmlReader xr)
        {            
            xr.Read();
            while(xr.IsEndElementWithName("metadataTypes")==false && xr.EOF==false)
            {
                if(xr.IsElementWithName("metadataType"))
                {
                    var item = new ExcelMetadataType(xr, _wb.RichData.IndexStore);
                    switch(item.Name)
                    {
                        case FUTURE_METADATA_DYNAMIC_ARRAY_NAME:
                            DynamicArrayTypeIndex = MetadataTypes.Count + 1;
                            break;
                        case FUTURE_METADATA_RICHDATA_NAME:
                            RichDataTypeIndex = MetadataTypes.Count + 1;
                            break;


                    }
                    MetadataTypes.Add(item);
                }
                xr.Read();
            }
        }

        internal int CreateDefaultXmlDynamicArray()
        {
            MetadataTypes.Add(new ExcelMetadataType(_wb.RichData.IndexStore) { Name = FUTURE_METADATA_DYNAMIC_ARRAY_NAME, MinSupportedVersion=120000, Flags=MetadataFlags.Copy | MetadataFlags.PasteAll | MetadataFlags.PasteValues | MetadataFlags.Merge | MetadataFlags.SplitFirst | MetadataFlags.RowColShift | MetadataFlags.ClearFormats | MetadataFlags.ClearComments | MetadataFlags.Assign | MetadataFlags.Coerce | MetadataFlags.CellMeta });
            var fmd = new FutureMetadataDynamicArray(true, _wb.RichData) { Name = FUTURE_METADATA_DYNAMIC_ARRAY_NAME };
            //fmd.Types.Add(new ExcelFutureMetadataDynamicArray(true));
            //FutureMetadata.Add(fmd.Name, fmd);
            FutureMetadata.Add(fmd);
            DynamicArrayTypeIndex = 1;

            var item = new ExcelCellMetadataBlock();
            item.Records.Add(new ExcelCellMetadataRecord(1, 0));
            CellMetadata.Add(item);
            return CellMetadata.Count;
        }

        //internal ExcelFutureMetadata GetFutureMetadataRichDataCollection()
        //{
        //    if (FutureMetadata.TryGetValue(FUTURE_METADATA_RICHDATA_NAME, out ExcelFutureMetadata fm))
        //    {
        //        return fm;
        //    }
        //    var mdt = new ExcelMetadataType(_wb.RichData.IndexStore) { Name = FUTURE_METADATA_RICHDATA_NAME, MinSupportedVersion = 120000, Flags = MetadataFlags.Copy | MetadataFlags.PasteAll | MetadataFlags.PasteValues | MetadataFlags.Merge | MetadataFlags.SplitFirst | MetadataFlags.RowColShift | MetadataFlags.ClearFormats | MetadataFlags.ClearComments | MetadataFlags.Assign | MetadataFlags.Coerce };
        //    MetadataTypes.Add(mdt);
        //    RichDataTypeIndex = MetadataTypes.Count;
        //    fm = new ExcelFutureMetadata() { Index = FutureMetadata.Count, Name = FUTURE_METADATA_RICHDATA_NAME };
        //    FutureMetadata.Add(FUTURE_METADATA_RICHDATA_NAME, fm);
        //    return fm;
        //}

        internal void CreateRichValueMetadata(ExcelRichData richData, int rvIndex, out int valueMetadataIndex)
        {
            if(!_metadataTypeNames.Contains(FUTURE_METADATA_RICHDATA_NAME))
            {
                var mdt = new ExcelMetadataType(_wb.RichData.IndexStore) { Name = FUTURE_METADATA_RICHDATA_NAME, MinSupportedVersion = 120000, Flags = MetadataFlags.Copy | MetadataFlags.PasteAll | MetadataFlags.PasteValues | MetadataFlags.Merge | MetadataFlags.SplitFirst | MetadataFlags.RowColShift | MetadataFlags.ClearFormats | MetadataFlags.ClearComments | MetadataFlags.Assign | MetadataFlags.Coerce };
                _metadataTypeNames.Add(FUTURE_METADATA_RICHDATA_NAME);
                RichDataTypeIndex = MetadataTypes.Count;
                MetadataTypes.Add(mdt);
            }
            if(FutureMetadataRichValue == null)
            {
                FutureMetadataRichValue = new FutureMetadataRichValue(FUTURE_METADATA_RICHDATA_NAME, _wb.RichData, this);
                MetadataTypes.CreateRelation(MetadataTypes[RichDataTypeIndex], FutureMetadataRichValue, IndexType.String);
                FutureMetadata.Add(FutureMetadataRichValue);
            }
            var block = new FutureMetadataRichDataBlock(_wb.RichData);
            var rel = richData.Values.CreateRelation(block, rvIndex, IndexType.ZeroBasedPointer);
            block.RichDataId = rel.To.Id;
            FutureMetadataRichValue.Blocks.Add(block);
            var mdItem = new ExcelValueMetadataBlock(this, RichDataTypeIndex + 1, rel.To.Id, _wb.RichData);
            ValueMetadata.Add(mdItem);
            valueMetadataIndex = ValueMetadata.Count;
            //var fmdRichDataCollection = GetFutureMetadataRichDataCollection();
            //var rdItem = new ExcelFutureMetadataRichData(richData.Values.Items.Count - 1);
            //fmdRichDataCollection.Types.Add(rdItem);
            //var mdItem = new ExcelValueMetadataBlock(this, _wb.RichData.IndexStore);
            //mdItem.Records.Add(new ExcelCellMetadataRecord(RichDataTypeIndex, fmdRichDataCollection.Types.Count - 1));
            //mdItem.CreateRelations();
            //ValueMetadata.Add(mdItem);
            //valueMetadataIndex = ValueMetadata.Count;
        }

        internal bool HasMetadata()
        {
            return MetadataTypes.Count==0;
        }

        internal bool IsFormulaDynamic(int cm)
        {
            if(cm <= CellMetadata.Count)
            {
                var cellMetadata = CellMetadata[cm - 1];
                var record = cellMetadata.Records.First();
                var metadataType = MetadataTypes[record.TypeIndex - 1];
                if (metadataType.Name == FUTURE_METADATA_DYNAMIC_ARRAY_NAME)
                {
                    return FutureMetadata[FUTURE_METADATA_DYNAMIC_ARRAY_NAME].Types[record.ValueIndex].IsDynamicArray;
                }
            }
            return false;
        }
        internal bool IsSpillError(int vm)
        {
            return GetErrorType(vm) == 8;
        }
        internal bool IsCalcError(int vm)
        {
            return GetErrorType(vm) == 13;
        }

        internal int GetErrorType(int vm)
        {
            if (ValueMetadata.Count >= vm) return -1;
            var valueMetadata = ValueMetadata[vm - 1];
            var record = valueMetadata.Records.First();
            var metadataType = MetadataTypes[record.TypeIndex - 1];
            if (metadataType.Name == FUTURE_METADATA_RICHDATA_NAME)
            {
                var ix = FutureMetadata[metadataType.Name].Types[record.ValueIndex].AsRichData.Index;
                var rd = _wb.RichData.Values.Items[ix];
                var erd = rd.As.Type<ErrorRichValueBase>();
                //var fieldIx = rd.Structure.Keys.FindIndex(x => x.Name == "errorType");
                if (erd != null && erd.ErrorType.HasValue)
                {
                    //return int.Parse(rd.Values[fieldIx]);
                    return erd.ErrorType.Value;
                }
            }
            return -1;
        }
        internal void GetDynamicArrayIndex(out int cm)
        {
            if(HasMetadata())
            {
                cm=CreateDefaultXmlDynamicArray();                
            }
            else
            {
                var tIx = FutureMetadata[FUTURE_METADATA_DYNAMIC_ARRAY_NAME].Index + 1;
                if (tIx >= 0)
                {
                    cm = CellMetadata.FindIndex(x => x.Records.Exists(y => y.TypeIndex == tIx)) + 1;
                    if(cm<=0)
                    {
                        var mtIx = MetadataTypes.FindIndex(x => x.Name == FUTURE_METADATA_DYNAMIC_ARRAY_NAME) + 1;
                        var item = new ExcelCellMetadataBlock();
                        item.Records.Add(new ExcelCellMetadataRecord(mtIx, tIx));
                        CellMetadata.Add(item);
                        cm = CellMetadata.Count;
                    }
                }
                else
                {
                    cm=CreateDefaultXmlDynamicArray();
                }
            }
        }

        internal void Save(ZipOutputStream stream, CompressionLevel compressionLevel, string fileName)
        {
            stream.PutNextEntry(fileName);
            stream.CompressionLevel = (OfficeOpenXml.Packaging.Ionic.Zlib.CompressionLevel)compressionLevel;
            var sw = new StreamWriter(stream);

            sw.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
            sw.Write($"<metadata xmlns=\"{Schemas.schemaMain}\" xmlns:xlrd=\"{Schemas.schemaRichData}\" xmlns:xda=\"{Schemas.schemaDynamicArray}\">");
            WriteMetadataTypes(sw);
            WriteMetadataStrings(sw);
            WriteMdxMetadata(sw);
            WriteFutureMetadata(sw);
            WriteCellMetadataItems(sw, "cellMetadata", CellMetadata);
            WriteValueMetadataItems(sw, "valueMetadata", ValueMetadata);
            sw.Write("</metadata>");
            sw.Flush();

        }

        internal void CreatePart()
        {
            if (_part == null)
            {
                _part = _wb._package.ZipPackage.CreatePart(_uri, ContentTypes.contentTypeMetaData);
                _wb.Part.CreateRelationship(_uri, TargetMode.Internal, Relationsships.schemaMetadata);
            }
            _part.SaveHandler = Save;
        }

        private void WriteValueMetadataItems(StreamWriter sw, string element, ValueMetadataBlockCollection collection)
        {
            if (collection.Count == 0) return;
            sw.Write($"<{element} count=\"{collection.Count}\">");
            foreach (var item in collection)
            {
                sw.Write("<bk>");
                foreach (var r in item.Records)
                {
                    sw.Write($"<rc t=\"{r.TypeIndex}\" v=\"{r.ValueIndex}\"/>");
                }
                sw.Write("</bk>");
            }
            sw.Write($"</{element}>");
        }

        private void WriteCellMetadataItems(StreamWriter sw, string element, List<ExcelCellMetadataBlock> collection)
        {
            if (collection.Count == 0) return;
            sw.Write($"<{element} count=\"{collection.Count}\">");
            foreach(var item in collection)
            {
                sw.Write("<bk>");
                foreach(var r in item.Records)
                {
                    sw.Write($"<rc t=\"{r.TypeIndex}\" v=\"{r.ValueIndex}\"/>");
                }
                sw.Write("</bk>");
            }
            sw.Write($"</{element}>");
        }
        private void WriteFutureMetadata(StreamWriter sw)
        {
            if (FutureMetadata.Count > 0)
            {
                foreach (var fmd in FutureMetadata.Values.OrderBy(x=>x.Index))
                {
                    sw.Write($"<futureMetadata name=\"{fmd.Name}\" count=\"1\">");
                    foreach(var t in fmd.Types)
                    {
                        sw.Write($"<bk><extLst><ext uri=\"{t.Uri}\">");
                        t.WriteXml(sw);
                        sw.Write($"</ext></extLst></bk>");
                    }
                    sw.Write($"</futureMetadata>");
                }
            }
        }

        private void WriteMetadataTypes(StreamWriter sw)
        {
            sw.Write($"<metadataTypes count=\"{MetadataTypes.Count}\">");
            foreach(var metadataType in MetadataTypes )
            {
                metadataType.WriteXml(sw);
            }
            sw.Write($"</metadataTypes>");
        }
        private void WriteMetadataStrings(StreamWriter sw)
        {
            if(!string.IsNullOrEmpty(_metadataStringsXml))
            {
                sw.Write($"<metadataStrings count=\"{_metadataStringCount}\">{_metadataStringsXml}</metadataStrings>");
            }
        }
        private void WriteMdxMetadata(StreamWriter sw)
        {
            if (!string.IsNullOrEmpty(_mdxMetadataXml))
            {
                sw.Write($"<mdxMetadata count=\"{_mdxMetadataCount}\">{_mdxMetadataXml}</metadataStrings>");
            }
        }

        internal bool IsDynamicArray(int cmIx)
        {
            var cm = CellMetadata[cmIx];            
            var t = MetadataTypes[cm.Records[0].TypeIndex-1];
            if(t.Name == FUTURE_METADATA_DYNAMIC_ARRAY_NAME)
            {
                if (FutureMetadata.TryGetValue(FUTURE_METADATA_DYNAMIC_ARRAY_NAME, out ExcelFutureMetadata fmd))
                {
                    var fmdt = fmd.Types[cm.Records[0].ValueIndex];
                    if (fmdt.Type==FutureMetadataType.DynamicArray)
                    {
                        return fmdt.IsDynamicArray;
                    }
                }
            }
            return false;
        }

        internal bool IsRichData(int vm)
        {
            if (vm > ValueMetadata.Count) return false;
            var valueMetadata = ValueMetadata[vm - 1];
            var t = MetadataTypes[valueMetadata.Records[0].TypeIndex - 1];
            if(t.Name == FUTURE_METADATA_RICHDATA_NAME)
            {
                
                if(FutureMetadata.TryGetValue(FUTURE_METADATA_RICHDATA_NAME, out ExcelFutureMetadata fmd))
                {
                    var fmdt = fmd.Types[valueMetadata.Records[0].ValueIndex];
                    if(fmdt.Type == FutureMetadataType.RichData)
                    {
                        return true;
                    }
                }
            }
            return false;
        }

    }
}