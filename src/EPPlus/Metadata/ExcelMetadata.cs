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
using OfficeOpenXml.RichData.RichValues;
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
        private readonly ExcelRichData _richData;
        private ZipPackagePart _part;
        private Uri _uri;
        
        //Preserve xml variables
        private string _metadataStringsXml;
        private string _metadataStringCount;
        private string _mdxMetadataXml;
        private string _mdxMetadataCount;
        public string _extLstXml;

        internal MetadataTypesCollection MetadataTypes { get; }
        //internal Dictionary<string, ExcelFutureMetadata> FutureMetadata { get; } = new Dictionary<string, ExcelFutureMetadata>();
        internal FutureMetadataCollection FutureMetadata { get; set; }

        internal FutureMetadataRichValue FutureMetadataRichValue { get; private set; }

        private readonly HashSet<string> _metadataTypeNames = new HashSet<string>();

        internal FutureMetadataDynamicArray FutureMetadataDynamicArray { get; private set; }
        internal List<ExcelCellMetadataBlock> CellMetadata { get; } = new List<ExcelCellMetadataBlock>();
        internal ValueMetadataBlockCollection ValueMetadata { get; }

        internal ValueMetadataRecordCollection ValueMetadataRecords { get; }

        internal FutureMetadataRichValueBlockCollection FutureMetadataBlocks { get; }
        internal int RichDataTypeIndex { get; private set; }
        internal int DynamicArrayTypeIndex { get; private set; }
        internal ZipPackagePart Part { get { return _part; } }

        public ExcelMetadata(ExcelWorkbook workbook)
        {
            //if(richData == null)
            //{
            //    richData = workbook.RichData;
            //}
            //_richData = richData;
            _wb = workbook;
            var p = _wb._package;
            ValueMetadata = new ValueMetadataBlockCollection(workbook.IndexStore);
            ValueMetadataRecords = new ValueMetadataRecordCollection(workbook.IndexStore);
            MetadataTypes = new MetadataTypesCollection(workbook.IndexStore);
            FutureMetadata = new FutureMetadataCollection(workbook.IndexStore);
            FutureMetadataBlocks = new FutureMetadataRichValueBlockCollection(workbook.IndexStore);
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
                        collection.Add(new ExcelValueMetadataBlock(xr, this, _wb.IndexStore));
                    }
                }
                xr.Read();
            }
        }

        private void ReadFutureMetadata(XmlReader xr)
        {
            //var item = new ExcelFutureMetadata();
            //item.Name = xr.GetAttribute("name");
            FutureMetadataBase fd;
            var name = xr.GetAttribute("name");
            if(name == FutureMetadataBase.DYNAMIC_ARRAY_NAME)
            {
                fd = new FutureMetadataDynamicArray(xr, _wb.IndexStore, this);
            }
            else if(name == FutureMetadataBase.RICHDATA_NAME)
            {
                fd = new FutureMetadataRichValue(xr, _wb.IndexStore, this);
            }
            else
            {
                fd = new FutureMetadataPreserve(xr, _wb.IndexStore);
            }
            fd.Index = FutureMetadata.Count;
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
                    var item = new ExcelMetadataType(xr, _wb.IndexStore);
                    switch(item.Name)
                    {
                        case FutureMetadataBase.DYNAMIC_ARRAY_NAME:
                            DynamicArrayTypeIndex = MetadataTypes.Count + 1;
                            break;
                        case FutureMetadataBase.RICHDATA_NAME:
                            RichDataTypeIndex = MetadataTypes.Count + 1;
                            break;


                    }
                    MetadataTypes.Add(item);
                }
                xr.Read();
            }
        }

        internal void InitRelations(ExcelRichData richData)
        {
            var fm = FutureMetadata.FirstOrDefault(x => x.Name == FutureMetadataBase.RICHDATA_NAME);
            var fmrv = fm as FutureMetadataRichValue;
            if(fmrv != null)
            {
                for(var ix = 0; ix < fmrv.Blocks.Count; ix++)
                {
                    var bk = fmrv.Blocks[ix];
                    bk.InitRelations(richData);
                }
            }
        }

        internal int CreateDefaultXmlDynamicArray()
        {
            MetadataTypes.Add(new ExcelMetadataType(_wb.IndexStore) { Name = FutureMetadataBase.DYNAMIC_ARRAY_NAME, MinSupportedVersion=120000, Flags=MetadataFlags.Copy | MetadataFlags.PasteAll | MetadataFlags.PasteValues | MetadataFlags.Merge | MetadataFlags.SplitFirst | MetadataFlags.RowColShift | MetadataFlags.ClearFormats | MetadataFlags.ClearComments | MetadataFlags.Assign | MetadataFlags.Coerce | MetadataFlags.CellMeta });
            //var fmd = new FutureMetadataDynamicArray(_wb.RichData) { Name = FUTURE_METADATA_DYNAMIC_ARRAY_NAME };
            //fmd.Types.Add(new ExcelFutureMetadataDynamicArray(true));
            //FutureMetadata.Add(fmd.Name, fmd);
            var fmd = FutureMetadataDynamicArray.GetDefault(_wb.IndexStore, this);
            DynamicArrayTypeIndex = FutureMetadata.Count;
            FutureMetadata.Add(fmd);

            var item = new ExcelCellMetadataBlock();
            item.Records.Add(new ExcelCellMetadataRecord(DynamicArrayTypeIndex - 1, 0));
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

        internal void CreateRichValueMetadata(ExcelRichData richData, ExcelRichValue richValue, out int valueMetadataIndex)
        {
            if(!_metadataTypeNames.Contains(FutureMetadataBase.RICHDATA_NAME))
            {
                var mdt = new ExcelMetadataType(_wb.IndexStore) { Name = FutureMetadataBase.RICHDATA_NAME, MinSupportedVersion = 120000, Flags = MetadataFlags.Copy | MetadataFlags.PasteAll | MetadataFlags.PasteValues | MetadataFlags.Merge | MetadataFlags.SplitFirst | MetadataFlags.RowColShift | MetadataFlags.ClearFormats | MetadataFlags.ClearComments | MetadataFlags.Assign | MetadataFlags.Coerce };
                
                _metadataTypeNames.Add(FutureMetadataBase.RICHDATA_NAME);
                RichDataTypeIndex = MetadataTypes.Count;
                MetadataTypes.Add(mdt);
            }
            var rdTypeId = MetadataTypes[RichDataTypeIndex].Id;
            if(FutureMetadataRichValue == null)
            {
                FutureMetadataRichValue = new FutureMetadataRichValue(FutureMetadataBase.RICHDATA_NAME, _wb.IndexStore, this);
                MetadataTypes.CreateRelation(MetadataTypes[RichDataTypeIndex], FutureMetadataRichValue, IndexType.String);
                FutureMetadata.Add(FutureMetadataRichValue);
            }
            var block = new FutureMetadataRichValueBlock(_wb.IndexStore);
            var rvIx = richData.Values.GetIndexById(richValue.Id);
            if(!rvIx.HasValue)
            {
                throw new ArgumentException("rich value was not added to RichData.Values");
            }
            FutureMetadataRichValue.Blocks.Add(block);
            var mdItem = new ExcelValueMetadataBlock(this, _wb.IndexStore);
            ValueMetadata.Add(mdItem);
            var rel = richData.Values.CreateRelation(block, rvIx.Value, IndexType.ZeroBasedPointer);
            block.RichDataId = rel.To.Id;
            mdItem.AddRecord(rdTypeId, rel.From.Id);
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
                if (metadataType.Name == FutureMetadataBase.DYNAMIC_ARRAY_NAME)
                {
                    var bk = FutureMetadata[FutureMetadataBase.DYNAMIC_ARRAY_NAME].Blocks[record.ValueIndex];
                    if(bk is FutureMetadataDynamicArrayBlock fmdab)
                    {
                        return fmdab.IsDynamicArray;
                    }
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
            var fmBk = record.GetFirstOutgoingSubRelation<FutureMetadataBlock>();
            if (fmBk != null)
            {
                var rv = fmBk.GetFirstTargetByType<ExcelRichValue>();
                var erd = rv.As.Type<ErrorRichValueBase>();
                if (erd != null && erd.ErrorType.HasValue)
                {
                    //return int.Parse(rd.Values[fieldIx]);
                    return erd.ErrorType.Value;
                }
            }
            //var metadataType = MetadataTypes[record.TypeIndex - 1];
            //if (metadataType.Name == FutureMetadataBase.RICHDATA_NAME)
            //{
            //    var rdId = FutureMetadata[metadataType.Name].Blocks[record.ValueIndex].FirstTargetId;
            //    if (!rdId.HasValue) return -1;
            //    var rd = _richData.Values.GetItem(rdId.Value);
            //    var erd = rd.As.Type<ErrorRichValueBase>();
            //    //var fieldIx = rd.Structure.Keys.FindIndex(x => x.Name == "errorType");
            //    if (erd != null && erd.ErrorType.HasValue)
            //    {
            //        //return int.Parse(rd.Values[fieldIx]);
            //        return erd.ErrorType.Value;
            //    }
            //}
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
                var tIx = FutureMetadata[FutureMetadataBase.DYNAMIC_ARRAY_NAME].Index + 1;
                if (tIx >= 0)
                {
                    cm = CellMetadata.FindIndex(x => x.Records.Exists(y => y.TypeIndex == tIx)) + 1;
                    if(cm<=0)
                    {
                        var mtIx = MetadataTypes.FindIndex(x => x.Name == FutureMetadataBase.DYNAMIC_ARRAY_NAME) + 1;
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
                    sw.Write($"<rc t=\"{r.MetadataTypeIndex}\" v=\"{r.FutureMetadataBlockIndex}\"/>");
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
                //foreach (var fmd in FutureMetadata.Values.OrderBy(x=>x.Index))
                //{
                //    sw.Write($"<futureMetadata name=\"{fmd.Name}\" count=\"1\">");
                //    foreach(var t in fmd.Types)
                //    {
                //        sw.Write($"<bk><extLst><ext uri=\"{t.Uri}\">");
                //        t.WriteXml(sw);
                //        sw.Write($"</ext></extLst></bk>");
                //    }
                //    sw.Write($"</futureMetadata>");
                //}
                foreach(var fmd in FutureMetadata)
                {
                    fmd.Save(sw);
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
            if(t.Name == FutureMetadataBase.DYNAMIC_ARRAY_NAME)
            {
                //if (FutureMetadata.TryGetValue(FUTURE_METADATA_DYNAMIC_ARRAY_NAME, out ExcelFutureMetadata fmd))
                //{
                //    var fmdt = fmd.Types[cm.Records[0].ValueIndex];
                //    if (fmdt.Type==FutureMetadataType.DynamicArray)
                //    {
                //        return fmdt.IsDynamicArray;
                //    }
                //}
                if(FutureMetadata.TryGetValue(FutureMetadataBase.DYNAMIC_ARRAY_NAME, out FutureMetadataBase fm))
                {
                    if (fm != null)
                    {
                        var vIx = cm.Records[0].ValueIndex;
                        var bk = fm.Blocks[vIx] as FutureMetadataDynamicArrayBlock;
                        if (bk != null) return bk.IsDynamicArray;
                    }
                }
                
            }
            return false;
        }

        internal bool IsRichData(int vm, out uint? richValueId)
        {
            richValueId = null;
            if (vm > ValueMetadata.Count) return false;
            var valueMetadata = ValueMetadata[vm - 1];
            var metadataType = valueMetadata.GetFirstOutgoingSubRelation<ExcelMetadataType>();
            if (metadataType == null || metadataType.Name != FutureMetadataBase.RICHDATA_NAME) return false;
            var futureMetadata = metadataType.GetFirstTargetByType<FutureMetadataBase>();
            if (futureMetadata == null) return false;
            var fmBlock = valueMetadata.GetFirstOutgoingSubRelation<FutureMetadataBlock>(out IndexRelation subRelation);
            if (fmBlock != null)
            {
                richValueId = subRelation.To.Id;
                return true;
            }
            //var t = MetadataTypes[valueMetadata.Records[0].TypeIndex - 1];
            //if(t.Name == FutureMetadataBase.RICHDATA_NAME)
            //{

            //    //if(FutureMetadata.TryGetValue(FUTURE_METADATA_RICHDATA_NAME, out ExcelFutureMetadata fmd))
            //    //{
            //    //    var fmdt = fmd.Types[valueMetadata.Records[0].ValueIndex];
            //    //    if(fmdt.Type == FutureMetadataType.RichData)
            //    //    {
            //    //        return true;
            //    //    }
            //    //}
            //    if (FutureMetadata.TryGetValue(FutureMetadataBase.RICHDATA_NAME, out FutureMetadataBase fm))
            //    {
            //        var vId = valueMetadata.FirstTargetId;
            //        if (vId.HasValue == false) return false;
            //        var bk = fm.Blocks.GetItem(vId.Value);
            //        return bk.Entity == RichDataEntities.FutureMetadataRichDataBlock;
            //    }

            //}
            return false;
        }

    }
}