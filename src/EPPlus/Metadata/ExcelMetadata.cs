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
        internal FutureMetadataCollection FutureMetadata { get; set; }

        internal FutureMetadataDynamicArray FutureMetadataDynamicArray { get; private set; }
        //internal List<ExcelCellMetadataBlock> CellMetadata { get; } = new List<ExcelCellMetadataBlock>();
        internal ValueMetadataBlockCollection ValueMetadata { get; }

        internal ValueMetadataRecordCollection ValueMetadataRecords { get; }

        internal CellMetadataBlockCollection CellMetadata { get; }

        internal CellMetadataRecordCollection CellMetadataRecords { get; set; }

        internal FutureMetadataRichValueBlockCollection FutureMetadataBlocks { get; }
        internal int DynamicArrayTypeIndex { get; private set; }
        internal ZipPackagePart Part { get { return _part; } }

        //public event EventHandler<ValueMetadataBlockDeletedEventArgs> ValueMetadataBlockDeleted;

        //public virtual void OnValueMetadataBlockDeleted(uint deletedEntityId)
        //{
        //    var args = new ValueMetadataBlockDeletedEventArgs(deletedEntityId);
        //    ValueMetadataBlockDeleted?.Invoke(this, args);
        //}

        internal EventHandler<ValueMetadataReadEventArgs> ValueMetadataRead;

        internal void OnValueMetadataRead(uint id, uint oneBasedIndex)
        {
            ValueMetadataRead?.Invoke(this, new ValueMetadataReadEventArgs(id, oneBasedIndex));
        }

        public ExcelMetadata(ExcelWorkbook workbook)
        {
            _wb = workbook;
            var p = _wb._package;
            ValueMetadata = new ValueMetadataBlockCollection(workbook.IndexStore);
            ValueMetadataRecords = new ValueMetadataRecordCollection(workbook.IndexStore);
            CellMetadata = new CellMetadataBlockCollection(workbook.IndexStore);
            CellMetadataRecords = new CellMetadataRecordCollection(workbook.IndexStore);
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

        private void ReadCellMetadataItems(XmlReader xr, string elementName, CellMetadataBlockCollection collection)
        {
            xr.Read();
            while(xr.IsEndElementWithName(elementName) ==false && xr.EOF==false)
            {
                if(xr.IsElementWithName("bk"))
                {
                    xr.Read();
                    while(xr.IsEndElementWithName("bk")==false)
                    {
                        collection.Add(new ExcelCellMetadataBlock(xr, this, _wb.IndexStore));
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
        }

        private void ReadMetadataTypes(XmlReader xr)
        {            
            xr.Read();
            while(xr.IsEndElementWithName("metadataTypes")==false && xr.EOF==false)
            {
                if(xr.IsElementWithName("metadataType"))
                {
                    var item = new ExcelMetadataType(xr, _wb.IndexStore);
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
                    fm.AddRelationTo(bk, IndexType.ZeroBasedPointer);
                }
            }
        }

        internal int CreateDefaultXmlDynamicArray()
        {
            var mt = new ExcelMetadataType(_wb.IndexStore) { Name = FutureMetadataBase.DYNAMIC_ARRAY_NAME, MinSupportedVersion = 120000, Flags = MetadataFlags.Copy | MetadataFlags.PasteAll | MetadataFlags.PasteValues | MetadataFlags.Merge | MetadataFlags.SplitFirst | MetadataFlags.RowColShift | MetadataFlags.ClearFormats | MetadataFlags.ClearComments | MetadataFlags.Assign | MetadataFlags.Coerce | MetadataFlags.CellMeta };
            MetadataTypes.Add(mt);
            var fmd = FutureMetadataDynamicArray.GetDefault(_wb.IndexStore, this, out uint bkId);
            DynamicArrayTypeIndex = FutureMetadata.Count;
            FutureMetadata.Add(fmd);

            var item = new ExcelCellMetadataBlock(_wb.Metadata, _wb.IndexStore);
            //item.Records.Add(new ExcelCellMetadataRecord(DynamicArrayTypeIndex - 1, 0));
            item.AddRecord(mt.Id, bkId);
            CellMetadata.Add(item);
            return CellMetadata.Count;
        }

        internal void CreateRichValueMetadata(ExcelRichData richData, ExcelRichValue richValue, out uint valueMetadataBlockId)
        {
            var rvMetadataType = default(ExcelMetadataType);
            if(MetadataTypes.TryGetValue(FutureMetadataBase.RICHDATA_NAME, out ExcelMetadataType mt))
            {
                rvMetadataType = mt;
            }
            else
            {
                rvMetadataType = new ExcelMetadataType(_wb.IndexStore) { Name = FutureMetadataBase.RICHDATA_NAME, MinSupportedVersion = 120000, Flags = MetadataFlags.Copy | MetadataFlags.PasteAll | MetadataFlags.PasteValues | MetadataFlags.Merge | MetadataFlags.SplitFirst | MetadataFlags.RowColShift | MetadataFlags.ClearFormats | MetadataFlags.ClearComments | MetadataFlags.Assign | MetadataFlags.Coerce };
                MetadataTypes.Add(rvMetadataType);
            }
            var rdTypeId = rvMetadataType.Id;
            var rvFutureMetadata = default(FutureMetadataBase);
            if(FutureMetadata.TryGetValue(FutureMetadataBase.RICHDATA_NAME, out FutureMetadataBase fmb))
            {
                rvFutureMetadata = fmb;
            }
            else
            {
                rvFutureMetadata = new FutureMetadataRichValue(FutureMetadataBase.RICHDATA_NAME, _wb.IndexStore, this);
                MetadataTypes.CreateRelation(rvMetadataType, rvFutureMetadata, IndexType.String);
                FutureMetadata.Add(rvFutureMetadata);
            }
            var block = new FutureMetadataRichValueBlock(_wb.IndexStore);
            rvFutureMetadata.Blocks.Add(block);
            rvFutureMetadata.AddRelationTo(block, IndexType.ZeroBasedPointer);
            var mdItem = new ExcelValueMetadataBlock(this, _wb.IndexStore);
            valueMetadataBlockId = mdItem.Id;
            ValueMetadata.Add(mdItem);
            var rel = block.AddRelationTo(richValue, IndexType.ZeroBasedPointer);
            block.RichDataId = rel.To.Id;
            mdItem.AddRecord(rdTypeId, rel.From.Id);
        }

        internal bool HasMetadata()
        {
            return MetadataTypes.Count==0;
        }

        internal bool IsFormulaDynamic(uint cmId)
        {
            var bk = CellMetadata.Get(cmId);
            if (bk != null)
            {
                if (bk.HasOutgoingRelationTo(RichDataEntities.MetadataType))
                {
                    var type = bk.GetFirstOutgoingRelByType<ExcelMetadataType>();
                    if (type != null && type.Name == FutureMetadataBase.DYNAMIC_ARRAY_NAME)
                    {
                        return true;
                    }
                }
            }
            return false;
        }
        //    if(cm <= CellMetadata.Count)
        //    {
        //        var cellMetadata = CellMetadata[cm - 1];
        //        var record = cellMetadata.Records.First();
        //        var metadataType = MetadataTypes[record.TypeIndex - 1];
        //        if (metadataType.Name == FutureMetadataBase.DYNAMIC_ARRAY_NAME)
        //        {
        //            var bk = FutureMetadata[FutureMetadataBase.DYNAMIC_ARRAY_NAME].Blocks[record.ValueIndex];
        //            if(bk is FutureMetadataDynamicArrayBlock fmdab)
        //            {
        //                return fmdab.IsDynamicArray;
        //            }
        //        }
        //    }
        //    return false;
        //}
        //internal bool IsSpillError(int vm)
        //{
        //    return GetErrorType(vm) == 8;
        //}
        //internal bool IsCalcError(int vm)
        //{
        //    return GetErrorType(vm) == 13;
        //}

        //internal int GetErrorType(int vm)
        //{
        //    if (ValueMetadata.Count >= vm) return -1;
        //    var valueMetadata = ValueMetadata[vm - 1];
        //    var record = valueMetadata.Records.First();
        //    var fmBk = record.GetFirstOutgoingSubRelation<FutureMetadataBlock>();
        //    if (fmBk != null)
        //    {
        //        var rv = fmBk.GetFirstOutgoingRelByType<ExcelRichValue>();
        //        var erd = rv.As.Type<ErrorRichValueBase>();
        //        if (erd != null && erd.ErrorType.HasValue)
        //        {
        //            return erd.ErrorType.Value;
        //        }
        //    }
        //    return -1;
        //}
        internal void GetDynamicArrayId(out uint cmId)
        {
            if(!MetadataTypes.TryGetValue(FutureMetadataBase.DYNAMIC_ARRAY_NAME, out ExcelMetadataType type))
            {
                CreateDefaultXmlDynamicArray();   
            }  
            cmId = FutureMetadata[FutureMetadataBase.DYNAMIC_ARRAY_NAME].Blocks[0].Id;
            //if(HasMetadata())
            //{
            //    cm=CreateDefaultXmlDynamicArray();                
            //}
            //else
            //{
            //    var tIx = FutureMetadata[FutureMetadataBase.DYNAMIC_ARRAY_NAME].Index + 1;
            //    if (tIx >= 0)
            //    {
            //        cm = FutureMetadata[FutureMetadataBase.DYNAMIC_ARRAY_NAME].Blocks.Get()
            //        //cm = CellMetadata.FindIndex(x => x.Records.Exists(y => y.TypeIndex == tIx)) + 1;
            //        if (cm<=0)
            //        {
            //            var mtIx = MetadataTypes.FindIndex(x => x.Name == FutureMetadataBase.DYNAMIC_ARRAY_NAME) + 1;
            //            var item = new ExcelCellMetadataBlock();
            //            item.Records.Add(new ExcelCellMetadataRecord(mtIx, tIx));
            //            CellMetadata.Add(item);
            //            cm = CellMetadata.Count;
            //        }
            //    }
            //    else
            //    {
            //        cm=CreateDefaultXmlDynamicArray();
            //    }
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
                if (item.Deleted) continue;
                var records = item.Records;
                if(records.Count() > 0)
                {
                    sw.Write("<bk>");
                    foreach (var r in records)
                    {
                        sw.Write($"<rc t=\"{r.MetadataTypeIndex}\" v=\"{r.FutureMetadataBlockIndex}\"/>");
                    }
                    sw.Write("</bk>");
                }
            }
            sw.Write($"</{element}>");
        }

        private void WriteCellMetadataItems(StreamWriter sw, string element, CellMetadataBlockCollection collection)
        {
            if (collection.Count == 0) return;
            sw.Write($"<{element} count=\"{collection.Count}\">");
            foreach(var item in collection)
            {
                sw.Write("<bk>");
                foreach(var r in item.Records)
                {
                    sw.Write($"<rc t=\"{r.MetadataTypeIndex}\" v=\"{r.FutureMetadataBlockIndex}\"/>");
                }
                sw.Write("</bk>");
            }
            sw.Write($"</{element}>");
        }
        private void WriteFutureMetadata(StreamWriter sw)
        {
            if (FutureMetadata.Count > 0)
            {
                foreach(var fmd in FutureMetadata)
                {
                    if(fmd.Deleted) continue;
                    fmd.Save(sw);
                }
            }
        }

        private void WriteMetadataTypes(StreamWriter sw)
        {
            if(MetadataTypes.Count > 0)
            {
                sw.Write($"<metadataTypes count=\"{MetadataTypes.Count}\">");
                foreach (var metadataType in MetadataTypes)
                {
                    if (metadataType.Deleted) continue;
                    metadataType.WriteXml(sw);
                }
                sw.Write($"</metadataTypes>");
            }
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

        internal bool IsDynamicIdByIndex(int index)
        {
            if (index >= CellMetadata.Count) return false;
            var cm = CellMetadata[index];
            var type = cm.GetFirstOutgoingRelByType<ExcelMetadataType>();
            if (type == null) return false;
            return type.Name == FutureMetadataBase.DYNAMIC_ARRAY_NAME;
        }

        internal bool IsDynamicArrayById(uint cmId)
        {
            var cm = CellMetadata.Get(cmId);
            if (cm == null) return false;
            var type = cm.GetFirstOutgoingRelByType<ExcelMetadataType>();
            if(type == null) return false;
            return type.Name == FutureMetadataBase.DYNAMIC_ARRAY_NAME;
            //var cm = CellMetadata[cmIx];            
            //var t = MetadataTypes[cm.Records[0].TypeIndex-1];
            //if(t.Name == FutureMetadataBase.DYNAMIC_ARRAY_NAME)
            //{
            //    if(FutureMetadata.TryGetValue(FutureMetadataBase.DYNAMIC_ARRAY_NAME, out FutureMetadataBase fm))
            //    {
            //        if (fm != null)
            //        {
            //            var vIx = cm.Records[0].ValueIndex;
            //            var bk = fm.Blocks[vIx] as FutureMetadataDynamicArrayBlock;
            //            if (bk != null) return bk.IsDynamicArray;
            //        }
            //    }
                
            //}
            //return false;
        }

        internal bool IsRichData(uint valueMetadataBlockId, out uint? richValueId)
        {
            richValueId = null;
            if (valueMetadataBlockId == 0) return false;
            var valueMetadata = ValueMetadata.Get(valueMetadataBlockId);
            if (valueMetadata == null)
            {
                return false;
            }
            var metadataType = valueMetadata.GetFirstOutgoingSubRelation<ExcelMetadataType>();
            if (metadataType == null || metadataType.Name != FutureMetadataBase.RICHDATA_NAME) return false;
            var futureMetadata = metadataType.GetFirstOutgoingRelByType<FutureMetadataBase>();
            if (futureMetadata == null) return false;
            var fmBlock = valueMetadata.GetFirstOutgoingSubRelation<FutureMetadataBlock>(out IndexRelation subRelation);
            if (fmBlock != null)
            {
                richValueId = subRelation.To.Id;
                return true;
            }
            return false;
        }

    }
}